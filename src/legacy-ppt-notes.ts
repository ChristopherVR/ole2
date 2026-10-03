/** Active MS-PPT notes: explicit slide/notes persist linkage and fixed text slots.
 * [MS-PPT] 2.4.14.6-7, 2.5.6/10/12 and PT_NotesBody (0x0C). */
import { readCompoundFileStream, replaceCompoundFileStream } from './ole2-stream-edit.js';
import { readRecordOrThrow, type PptRecord } from './legacy-ppt-record-stream.js';
import { RT, OA, HEADER_TOKEN_ENCRYPTED, HEADER_TOKEN_PLAIN } from './legacy-ppt-record-types.js';
import { buildPersistDirectory, parseUserEditAtom } from './ppt/persist-directory.js';
import { PptTextError, readPptSlideTexts, type PptTextAtom } from './legacy-ppt-text.js';
import { readPptSlideShapes } from './legacy-ppt-shape-reader.js';

export interface PptNotesReadLimits { maxRecords?: number; maxNotes?: number; maxTextBytes?: number }
export interface PptNoteTextAtom extends PptTextAtom {
  shapeId: number; shapeHeaderOffset: number; placeholderId?: number;
  role: 'body' | 'other'; editRefusal?: string;
}
export interface PptSlideNotes {
  slideId: number; slidePersistId: number; notesId: number; persistId: number;
  headerOffset: number; texts: PptNoteTextAtom[];
}
const viewOf = (bytes: Uint8Array) => new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
function fail(message: string): never { throw new PptTextError('corrupt', message); }
function inspect(input: Uint8Array, limits: PptNotesReadLimits = {}): { stream: Uint8Array; notes: PptSlideNotes[] } {
  const maxRecords = limits.maxRecords ?? 100000, maxNotes = limits.maxNotes ?? 10000, maxTextBytes = limits.maxTextBytes ?? 16 * 1024 * 1024;
  for (const value of [maxRecords, maxNotes, maxTextBytes]) if (!Number.isSafeInteger(value) || value < 1) fail('Invalid notes resource limit');
  const stream = readCompoundFileStream(input, ['PowerPoint Document']), user = readCompoundFileStream(input, ['Current User']);
  if (!stream || !user) fail('Missing PowerPoint streams');
  const view = viewOf(stream), uv = viewOf(user);
  const record = (offset: number, type?: number): PptRecord => {
    const r = readRecordOrThrow(view, offset);
    if (r.dataOffset + r.recLen > view.byteLength || (type !== undefined && r.recType !== type)) fail('Invalid notes record bounds/type');
    return r;
  };
  const cu = readRecordOrThrow(uv, 0);
  if (cu.recType !== RT.CurrentUserAtom || cu.recVer !== 0 || cu.recInstance !== 0 || cu.recLen < 24 || cu.dataOffset + cu.recLen > uv.byteLength) fail('Invalid CurrentUserAtom');
  const token = uv.getUint32(cu.dataOffset + 4, true);
  if (token === HEADER_TOKEN_ENCRYPTED) throw new PptTextError('encrypted', 'Encrypted PPT notes are unsupported');
  if (token !== HEADER_TOKEN_PLAIN) fail('Unknown encryption token');
  const { currentEdit, directory } = buildPersistDirectory(view, uv.getUint32(cu.dataOffset + 8, true), { maxEntries: maxRecords });
  if (currentEdit.encryptSessionPersistIdRef !== undefined) throw new PptTextError('encrypted', 'Encrypted notes persist objects are unsupported');
  let recordCount = 0, textBytes = 0;
  const children = (r: PptRecord): PptRecord[] => {
    const result: PptRecord[] = [], end = r.dataOffset + r.recLen;
    for (let pos = r.dataOffset; pos < end;) {
      if (++recordCount > maxRecords) fail('Notes record budget exceeded');
      if (end - pos < 8) fail('Truncated notes child header');
      const child = record(pos);
      if (child.dataOffset + child.recLen > end) fail('Notes child exceeds parent');
      result.push(child); pos = child.dataOffset + child.recLen;
    }
    return result;
  };
  const docOffset = directory.get(currentEdit.docPersistIdRef);
  if (docOffset === undefined) fail('Missing active document');
  const doc = record(docOffset, RT.Document);
  if (doc.recVer !== 15) fail('Invalid document container');
  const docChildren = children(doc), refs = new Map<number, number>(), persistIds = new Set<number>();
  const notesLists = docChildren.filter(r => r.recType === RT.SlideListWithText && r.recInstance === 2);
  if (notesLists.length > 1) fail('Ambiguous notes list');
  for (const list of notesLists) {
    if (list.recVer !== 15) fail('Invalid notes list');
    for (const r of children(list)) {
      if (r.recType !== RT.SlidePersistAtom || r.recVer !== 0 || r.recInstance !== 0 || r.recLen !== 20) fail('Invalid NotesPersistAtom');
      const persistId = view.getUint32(r.dataOffset, true), notesId = view.getUint32(r.dataOffset + 12, true);
      if (!notesId || !persistId || refs.has(notesId) || persistIds.has(persistId)) fail('Duplicate or null notes identity');
      if (refs.size >= maxNotes) fail('Notes count budget exceeded');
      refs.set(notesId, persistId); persistIds.add(persistId);
    }
  }
  // The shape reader preflights active slide record/text budgets before its
  // text decoder, including for standalone notes reader calls.
  const slides = readPptSlideShapes(input, { maxRecords, maxTextBytes, maxShapes: maxRecords });
  if (slides.length > maxNotes) fail('Active slide budget exceeded');
  const links = new Map<number, number>(), linkedNotes = new Set<number>();
  for (const slide of slides) {
    const offset = directory.get(slide.persistId);
    if (offset === undefined) fail('Missing active slide');
    const sr = record(offset, RT.Slide), slideAtoms = children(sr).filter(r => r.recType === RT.SlideAtom);
    if (slideAtoms.length !== 1 || slideAtoms[0]!.recVer !== 2 || slideAtoms[0]!.recInstance !== 0 || slideAtoms[0]!.recLen !== 24) fail('Invalid SlideAtom notes linkage');
    const notesId = view.getUint32(slideAtoms[0]!.dataOffset + 16, true);
    if (notesId && linkedNotes.has(notesId)) fail('Notes linked by multiple active slides');
    if (notesId) linkedNotes.add(notesId);
    links.set(slide.slideId, notesId);
  }
  const activeNotes = new Set(links.values());
  // Preflight linked notes before decoding notes strings. Unreferenced objects
  // and prior save-history bytes remain opaque.
  const raw = new Map<number, { header: PptRecord; backref: number; records: PptRecord[] }>();
  for (const [notesId, persistId] of refs) {
    if (!activeNotes.has(notesId)) continue;
    const offset = directory.get(persistId);
    if (offset === undefined) fail('Missing notes persist object');
    const header = record(offset, RT.Notes);
    if (header.recVer !== 15 || header.recInstance !== 0) fail('Invalid NotesContainer');
    const direct = children(header), atoms = direct.filter(r => r.recType === RT.NotesAtom);
    if (atoms.length !== 1 || atoms[0]!.recVer !== 1 || atoms[0]!.recInstance !== 0 || atoms[0]!.recLen !== 8) fail('Invalid NotesAtom');
    const records: PptRecord[] = [...direct], stack = direct.filter(r => r.recVer === 15 || r.recType === OA.ClientTextbox).map(r => ({ r, depth: 1 }));
    while (stack.length) {
      const item = stack.pop()!;
      if (item.depth > 128) fail('Notes nesting budget exceeded');
      for (const r of children(item.r)) {
        records.push(r);
        if (r.recType === RT.TextCharsAtom || r.recType === RT.TextBytesAtom) {
          textBytes += r.recLen;
          if (textBytes > maxTextBytes) fail('Notes text byte budget exceeded');
          if (r.recVer !== 0 || r.recInstance !== 0 || (r.recType === RT.TextCharsAtom && r.recLen % 2)) fail('Invalid notes text atom');
        }
        if (r.recVer === 15 || r.recType === OA.ClientTextbox) stack.push({ r, depth: item.depth + 1 });
      }
    }
    raw.set(notesId, { header, backref: view.getUint32(atoms[0]!.dataOffset, true), records });
  }
  const notes: PptSlideNotes[] = [];
  for (const slide of slides) {
    const notesId = links.get(slide.slideId)!;
    if (!notesId) continue;
    const info = raw.get(notesId), persistId = refs.get(notesId);
    if (!info || persistId === undefined || info.backref !== slide.slideId) fail('Notes ID or slide back-reference mismatch');
    const all = info.records;
    let notesMirror = false;
    for (const r of all.filter(r => [OA.FOPT, OA.TertiaryFOPT, 0xf121].includes(r.recType))) {
      if (r.recVer !== 3 || r.recInstance * 6 > r.recLen) fail('Invalid notes property table');
      for (let i = 0; i < r.recInstance; i++) if ((view.getUint16(r.dataOffset + i * 6, true) & 0x3fff) === 0x3a9) notesMirror = true;
    }
    const texts: PptNoteTextAtom[] = [], shapeIds = new Set<number>();
    const shapeRecords: Array<{ shape: PptRecord; within: PptRecord[] }> = [];
    let activeShape: typeof shapeRecords[number] | undefined;
    for (const r of all.sort((a, b) => a.headerOffset - b.headerOffset)) {
      if (activeShape && r.headerOffset >= activeShape.shape.dataOffset + activeShape.shape.recLen) activeShape = undefined;
      if (r.recType === OA.SpContainer) {
        if (activeShape) fail('Nested notes shape containers');
        activeShape = { shape: r, within: [] }; shapeRecords.push(activeShape);
      } else if (activeShape) activeShape.within.push(r);
    }
    for (const { shape, within } of shapeRecords) {
      const fsps = within.filter(r => r.recType === OA.FSP);
      if (fsps.length !== 1 || fsps[0]!.recVer !== 2 || fsps[0]!.recLen !== 8) fail('Invalid notes shape identity');
      const shapeId = view.getUint32(fsps[0]!.dataOffset, true);
      if (!shapeId || shapeIds.has(shapeId)) fail('Duplicate notes shape identity');
      shapeIds.add(shapeId);
      const placeholders = within.filter(r => r.recType === RT.OEPlaceholderAtom);
      if (placeholders.some(r => r.recVer !== 0 || r.recInstance !== 0 || r.recLen !== 8)) fail('Invalid notes PlaceholderAtom');
      const placeholderId = placeholders.length === 1 ? view.getUint8(placeholders[0]!.dataOffset + 4) : undefined;
      const atoms = within.filter(r => r.recType === RT.TextCharsAtom || r.recType === RT.TextBytesAtom);
      const headers = within.filter(r => r.recType === RT.TextHeaderAtom);
      const textboxes = within.filter(r => r.recType === OA.ClientTextbox);
      const body = placeholderId === 12 && headers.length === 1 && headers[0]!.recVer === 0 && headers[0]!.recInstance === 0 && headers[0]!.recLen === 4 && view.getUint32(headers[0]!.dataOffset, true) === 2;
      for (const atom of atoms) {
        const wide = atom.recType === RT.TextCharsAtom;
        let text = '';
        for (let p = atom.dataOffset; p < atom.dataOffset + atom.recLen; p += wide ? 2 : 1) {
          const value = wide ? view.getUint16(p, true) : view.getUint8(p);
          if (!value) fail('NUL in notes text atom');
          text += String.fromCharCode(value);
        }
        texts.push({ shapeId, shapeHeaderOffset: shape.headerOffset, placeholderId, role: body ? 'body' : 'other', text, headerOffset: atom.headerOffset,
          encoding: wide ? 'utf16' : 'compressed-unicode', editRefusal: notesMirror ? 'Notes have an OOXML mirror' : !body ? 'Not a validated notes body placeholder' : atoms.length !== 1 || placeholders.length !== 1 || textboxes.length !== 1 || headers[0]!.headerOffset < textboxes[0]!.dataOffset || headers[0]!.headerOffset > atom.headerOffset || atom.dataOffset + atom.recLen > textboxes[0]!.dataOffset + textboxes[0]!.recLen ? 'Ambiguous notes text slot' : undefined });
      }
    }
    notes.push({ slideId: slide.slideId, slidePersistId: slide.persistId, notesId, persistId, headerOffset: info.header.headerOffset, texts });
  }
  return { stream, notes };
}
export function readPptSlideNotes(input: Uint8Array, limits: PptNotesReadLimits = {}): PptSlideNotes[] { return inspect(input, limits).notes; }
/** Notes cannot own bytes simultaneously owned by another persist object,
 * including opaque slide content and save-chain metadata. */
function ownershipRefusal(input: Uint8Array, stream: Uint8Array, note: PptSlideNotes): string | undefined {
  const view = viewOf(stream), user = readCompoundFileStream(input, ['Current User'])!;
  const uv = viewOf(user), cu = readRecordOrThrow(uv, 0), editOffset = uv.getUint32(cu.dataOffset + 8, true);
  const { directory } = buildPersistDirectory(view, editOffset);
  const owned = readRecordOrThrow(view, note.headerOffset), start = owned.headerOffset, end = owned.dataOffset + owned.recLen;
  const overlaps = (offset: number): boolean => {
    const other = readRecordOrThrow(view, offset), otherEnd = other.dataOffset + other.recLen;
    if (otherEnd > view.byteLength) fail('Unbounded live persist object prevents safe notes edit');
    return start < otherEnd && other.headerOffset < end;
  };
  for (const [persistId, offset] of directory) {
    if (persistId !== note.persistId && overlaps(offset)) return 'Notes container overlaps another live persist object';
  }
  let offset = editOffset;
  const visited = new Set<number>();
  while (offset) {
    if (visited.has(offset) || visited.size >= 10000) return 'Invalid notes save-history ownership';
    visited.add(offset);
    const edit = parseUserEditAtom(view, offset);
    if (overlaps(offset) || overlaps(edit.offsetPersistDirectory)) return 'Notes container overlaps save-history metadata';
    offset = edit.offsetLastEdit;
  }
  return undefined;
}
export type PptNotesEditResult = { status: 'edited' | 'unchanged'; bytes: Uint8Array } | { status: 'unsupported'; bytes: Uint8Array; reason: string };
export function editPptNotesText(input: Uint8Array, edit: {
  slideId: number; slidePersistId: number; notesId: number; persistId: number; shapeId: number;
  expectedHeaderOffset: number; expectedText: string; text: string;
}): PptNotesEditResult {
  const reject = (reason: string): PptNotesEditResult => ({ status: 'unsupported', bytes: input, reason });
  try {
    // Snapshot caller accessors once before parsing/validation/encoding.
    const { slideId, slidePersistId, notesId, persistId, shapeId, expectedHeaderOffset, expectedText, text } = edit;
    if (typeof text !== 'string' || typeof expectedText !== 'string') return reject('Notes replacement must be a string');
    const { stream, notes } = inspect(input);
    const selected = notes.filter(n => n.slideId === slideId && n.slidePersistId === slidePersistId && n.notesId === notesId && n.persistId === persistId);
    const matches = selected.length === 1 ? selected[0]!.texts.filter(t => t.shapeId === shapeId && t.headerOffset === expectedHeaderOffset) : [];
    const atom = matches[0];
    if (matches.length !== 1 || !atom || atom.text !== expectedText) return reject('Notes identity or expected text no longer matches');
    if (text === atom.text) return { status: 'unchanged', bytes: input };
    if (atom.editRefusal) return reject(atom.editRefusal);
    const ownership = ownershipRefusal(input, stream, selected[0]!);
    if (ownership) return reject(ownership);
    if (notes.reduce((n, note) => n + note.texts.filter(t => t.headerOffset === atom.headerOffset).length, 0) !== 1) return reject('Notes text atom is shared by active locations');
    const start = atom.headerOffset, end = start + 8 + atom.text.length * (atom.encoding === 'utf16' ? 2 : 1);
    for (const slide of readPptSlideTexts(input).slides) for (const text of slide.texts) {
      const textEnd = text.headerOffset + 8 + text.text.length * (text.encoding === 'utf16' ? 2 : 1);
      if (start < textEnd && text.headerOffset < end) return reject('Notes slot overlaps active slide text');
    }
    if (text.length !== atom.text.length) return reject('Notes replacement must keep the UTF-16 character count');
    for (let i = 0; i < text.length; i++) {
      const old = atom.text.charCodeAt(i), value = text.charCodeAt(i);
      if (!value || ((old < 32 || value < 32 || old === 42 || value === 42) && old !== value)) return reject('Notes controls and field markers must remain at the same offsets');
      if (atom.encoding === 'compressed-unicode' && value > 255) return reject('Notes character does not fit existing encoding');
      if ((value >= 0xd800 && value <= 0xdbff && !(text.charCodeAt(i + 1) >= 0xdc00 && text.charCodeAt(i + 1) <= 0xdfff)) ||
          (value >= 0xdc00 && value <= 0xdfff && !(text.charCodeAt(i - 1) >= 0xd800 && text.charCodeAt(i - 1) <= 0xdbff))) return reject('Unpaired notes UTF-16 surrogate');
    }
    const next = stream.slice(), view = viewOf(next);
    for (let i = 0; i < text.length; i++) {
      if (atom.encoding === 'utf16') view.setUint16(atom.headerOffset + 8 + i * 2, text.charCodeAt(i), true);
      else next[atom.headerOffset + 8 + i] = text.charCodeAt(i);
    }
    const bytes = replaceCompoundFileStream(input, ['PowerPoint Document'], next);
    if (bytes === input) return reject('Notes compound stream cannot be safely replaced');
    const candidate = inspect(bytes).notes.find(n => n.slideId === slideId && n.notesId === notesId && n.persistId === persistId);
    const confirmed = candidate?.texts.filter(t => t.shapeId === shapeId && t.headerOffset === expectedHeaderOffset);
    if (confirmed?.length !== 1 || confirmed[0]!.text !== text) return reject('Candidate notes revalidation failed');
    return { status: 'edited', bytes };
  } catch (error) { return reject(error instanceof Error ? error.message : 'Malformed PPT notes'); }
}
