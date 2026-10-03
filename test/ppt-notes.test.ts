import { readFileSync } from 'node:fs';
import { describe, expect, it } from 'vitest';
import { PptDocument } from '../src/ppt-document.js';
import { UnsupportedOle2EditError } from '../src/ole2-document-base.js';
import { readPptSlideNotes, editPptNotesText } from '../src/legacy-ppt-notes.js';
import { readCompoundFileStream, replaceCompoundFileStream } from '../src/ole2-stream-edit.js';
import { pptShapeChildren } from '../src/legacy-ppt-shape-reader.js';
import { readRecordOrThrow } from '../src/legacy-ppt-record-stream.js';
import { RT, OA } from '../src/legacy-ppt-record-types.js';
import { buildPersistDirectory } from '../src/ppt/persist-directory.js';
import { resizeCompoundFileStream } from '../src/ole2-stream-resize.js';

const load = () => new Uint8Array(readFileSync(new URL('./fixtures/ppt/native-text.ppt', import.meta.url)));
const viewOf = (b: Uint8Array) => new DataView(b.buffer, b.byteOffset, b.byteLength);
const body = (d: PptDocument, index = 0) => d.slides[index]!.notes!.texts.find(t => t.role === 'body')!;
function mutate(fn: (view: DataView, notes: ReturnType<typeof readPptSlideNotes>) => void): Uint8Array {
  const input = load(), notes = readPptSlideNotes(input), stream = readCompoundFileStream(input, ['PowerPoint Document'])!;
  fn(viewOf(stream), notes); return replaceCompoundFileStream(input, ['PowerPoint Document'], stream);
}
function editFor(input: Uint8Array) {
  const note = readPptSlideNotes(input)[0]!, atom = note.texts.find(t => t.role === 'body')!;
  return { slideId: note.slideId, slidePersistId: note.slidePersistId, notesId: note.notesId, persistId: note.persistId,
    shapeId: atom.shapeId, expectedHeaderOffset: atom.headerOffset, expectedText: atom.text, text: 'Synthetic notes modified.' };
}
describe('active binary PPT notes model', () => {
  it('resolves notes ID/persist/back-reference and body placeholders without treating field text as notes body', () => {
    const d = new PptDocument(load());
    expect(d.slides.map(s => [s.notes!.slideId, s.notes!.slidePersistId, s.notes!.notesId, s.notes!.persistId])).toEqual([[256, 15, 256, 17], [257, 16, 257, 18]]);
    expect(body(d).text).toBe('Synthetic notes retained.'); expect(body(d, 1).text).toBe('Unicode notes retained.');
    expect(body(d).placeholderId).toBe(12); expect(body(d).encoding).toBe('compressed-unicode');
    expect(d.slides[0]!.notes!.texts.find(t => t.text === '*')!.role).toBe('other');
    expect(d.serialize()).toEqual(load()); expect(d.dirty).toBe(false);
    expect(Object.isFrozen(d.slides[0]!.notes)).toBe(true); expect(Object.isFrozen(d.slides[0]!.notes!.texts)).toBe(true);
    expect(Reflect.set(d.slides[0]!, 'notes', undefined)).toBe(false);
  });
  it('writes only existing body atom bytes, preserves all opaque bytes, refreshes handles through multiple edits', () => {
    const input = load(), copy = input.slice(), d = new PptDocument(input), retained = body(d), original = retained.text;
    retained.text = 'Synthetic notes modified.';
    expect(retained.text).toBe('Synthetic notes modified.'); expect(d.revision).toBe(1);
    const after = d.serialize(), a = readCompoundFileStream(input, ['PowerPoint Document'])!, b = readCompoundFileStream(after, ['PowerPoint Document'])!;
    const off = retained.headerOffset + 8;
    expect(b.subarray(0, off)).toEqual(a.subarray(0, off)); expect(b.subarray(off + original.length)).toEqual(a.subarray(off + original.length));
    body(d, 1).text = 'Unicode notes modified.'; d.slides[0]!.shapes[0]!.x += 8;
    d.slides[0]!.texts[0]!.text = 'Native title updated'; retained.text = original;
    expect(retained.text).toBe(original); expect(d.revision).toBe(5); expect(input).toEqual(copy);
    const reopened = new PptDocument(d.serialize()); expect(body(reopened, 1).text).toBe('Unicode notes modified.'); expect(reopened.slides[0]!.shapes[0]!.x).toBe(248);
  });
  it('keeps body and field no-ops clean and explicitly refuses field mutation', () => {
    const d = new PptDocument(load()), field = d.slides[0]!.notes!.texts.find(t => t.role === 'other')!;
    body(d).text = body(d).text; field.text = field.text;
    expect(d.dirty).toBe(false); expect(d.revision).toBe(0);
    expect(() => { field.text = '+'; }).toThrow(/validated notes body/); expect(d.dirty).toBe(false);
    expect(() => body(d).runs).toThrow(UnsupportedOle2EditError);
  });
  it.each(['This replacement grows notes.', 'Synthetic\nnotes retained.', 'Synthetic notes retaine\0.', 'Synthetic notes retaineΩ.', '*ynthetic notes retained.'])('refuses unsafe replacement %j atomically', text => {
    const input = load(), d = new PptDocument(input), t = body(d), original = t.text;
    expect(() => { t.text = text; }).toThrow(UnsupportedOle2EditError);
    expect(t.text).toBe(original); expect(d.serialize()).toEqual(input); expect(d.revision).toBe(0);
  });
  it.each(['notesId', 'persistId', 'slideId', 'slidePersistId', 'shapeId', 'expectedHeaderOffset'] as const)('refuses stale %s low-level identity', key => {
    const input = load(), e = editFor(input); e[key] += 1;
    const result = editPptNotesText(input, e); expect(result.status).toBe('unsupported'); expect(result.bytes).toBe(input);
  });
  it('snapshots edit accessor text exactly once before validation and byte encoding', () => {
    const input = load(); let reads = 0;
    const result = editPptNotesText(input, { ...editFor(input), get text() { reads++; return reads === 1 ? 'Synthetic notes modified.' : 'bad'; } });
    expect(reads).toBe(1); expect(result.status).toBe('edited'); expect(readPptSlideNotes(result.bytes)[0]!.texts[0]!.text).toBe('Synthetic notes modified.');
  });
  it('rejects retained text handles when physical shape identity changes', () => {
    class Adopting extends PptDocument { adopt(bytes: Uint8Array) { this.commitBytes(bytes); } }
    const d = new Adopting(load()), t = body(d);
    const changed = mutate((v, notes) => {
      const shape = readRecordOrThrow(v, notes[0]!.texts[0]!.shapeHeaderOffset), fsp = pptShapeChildren(v, shape).find(r => r.recType === OA.FSP)!;
      v.setUint32(fsp.dataOffset, v.getUint32(fsp.dataOffset, true) + 7, true);
    });
    d.adopt(changed); const revision = d.revision;
    expect(() => t.text).toThrow(/identity/); expect(() => { t.text = 'Synthetic notes modified.'; }).toThrow(/identity/);
    expect(d.revision).toBe(revision); expect(d.serialize()).toEqual(changed);
  });
  it('refuses notes with mirrors without changing getters or bytes', () => {
    const input = mutate((v, notes) => {
      const shape = readRecordOrThrow(v, notes[0]!.texts[0]!.shapeHeaderOffset), table = pptShapeChildren(v, shape).find(r => r.recType === OA.FOPT)!;
      v.setUint16(table.dataOffset, 0x3a9, true);
    });
    const d = new PptDocument(input), t = body(d);
    expect(t.editRefusal).toMatch(/mirror/); expect(() => { t.text = 'Synthetic notes modified.'; }).toThrow(/mirror/);
    expect(d.serialize()).toEqual(input); expect(d.revision).toBe(0);
  });
  it('diagnoses malformed notes without blocking existing slide reading or claiming notes absent', () => {
    const input = mutate((v, notes) => { const a = pptShapeChildren(v, readRecordOrThrow(v, notes[0]!.headerOffset)).find(r => r.recType === RT.NotesAtom)!; v.setUint32(a.dataOffset, 999, true); });
    const d = new PptDocument(input);
    expect(d.slides[0]!.notes).toBeUndefined(); expect(d.slides[0]!.notesStatus).toBe('unsupported');
    expect(d.slides[0]!.notesDiagnostic).toMatch(/back-reference/); expect(d.unsupported).toContain('notes-decoding');
    expect(d.slides[0]!.texts[0]!.text).toBe('Native title fixture'); expect(d.serialize()).toEqual(input);
    expect(() => readPptSlideNotes(input)).toThrow(/back-reference/);
    expect(editPptNotesText(input, editFor(load())).status).toBe('unsupported');
  });
  it('rejects a notes page linked by multiple active slides', () => {
    const input = mutate((v, notes) => {
      const user = readCompoundFileStream(load(), ['Current User'])!, uv = viewOf(user);
      const chain = buildPersistDirectory(v, uv.getUint32(16, true));
      const slide = readRecordOrThrow(v, chain.directory.get(notes[1]!.slidePersistId)!);
      const atom = pptShapeChildren(v, slide).find(r => r.recType === RT.SlideAtom)!;
      v.setUint32(atom.dataOffset + 16, notes[0]!.notesId, true);
    });
    expect(() => readPptSlideNotes(input)).toThrow(/multiple active slides/);
    expect(new PptDocument(input).slides[0]!.notesDiagnostic).toMatch(/multiple active slides/);
  });
  it('leaves unreferenced notes objects opaque, including unsupported/malformed stale record data', () => {
    const input = mutate((v, notes) => {
      const user = readCompoundFileStream(load(), ['Current User'])!, uv = viewOf(user), chain = buildPersistDirectory(v, uv.getUint32(16, true));
      const slide = readRecordOrThrow(v, chain.directory.get(notes[1]!.slidePersistId)!);
      const atom = pptShapeChildren(v, slide).find(r => r.recType === RT.SlideAtom)!;
      v.setUint32(atom.dataOffset + 16, 0, true); v.setUint16(notes[1]!.headerOffset + 8, 0, true);
    });
    const d = new PptDocument(input); expect(d.slides[1]!.notes).toBeUndefined(); body(d).text = 'Synthetic notes modified.';
    const a = readCompoundFileStream(input, ['PowerPoint Document'])!, b = readCompoundFileStream(d.serialize(), ['PowerPoint Document'])!;
    const offset = readPptSlideNotes(load())[1]!.headerOffset;
    expect(b.subarray(offset)).toEqual(a.subarray(offset));
  });
  it('refuses notes nested in opaque active-slide payload even when slide text does not expose the alias', () => {
    const input = load(), stream = readCompoundFileStream(input, ['PowerPoint Document'])!, user = readCompoundFileStream(input, ['Current User'])!;
    const view = viewOf(stream), chain = buildPersistDirectory(view, viewOf(user).getUint32(16, true));
    const notes = readPptSlideNotes(input)[0]!, nr = readRecordOrThrow(view, notes.headerOffset), sr = readRecordOrThrow(view, chain.directory.get(notes.slidePersistId)!);
    const noteBytes = stream.subarray(nr.headerOffset, nr.dataOffset + nr.recLen), slideBytes = stream.subarray(sr.headerOffset, sr.dataOffset + sr.recLen);
    const oldEditOffset = viewOf(user).getUint32(16, true), oldEdit = readRecordOrThrow(view, oldEditOffset);
    const oldEditBytes = stream.subarray(oldEdit.headerOffset, oldEdit.dataOffset + oldEdit.recLen);
    const directoryOffset = stream.length + slideBytes.length + 8 + noteBytes.length, newEditOffset = directoryOffset + 24;
    const appended = new Uint8Array(newEditOffset + oldEditBytes.length), av = viewOf(appended);
    appended.set(stream); appended.set(slideBytes, stream.length);
    av.setUint32(stream.length + 4, sr.recLen + 8 + noteBytes.length, true);
    const wrapper = stream.length + slideBytes.length;
    av.setUint16(wrapper + 2, 0x7abc, true); av.setUint32(wrapper + 4, noteBytes.length, true); appended.set(noteBytes, wrapper + 8);
    // Append a proper new save so each new object precedes its directory.
    av.setUint16(directoryOffset + 2, RT.PersistDirectoryAtom, true); av.setUint32(directoryOffset + 4, 16, true);
    av.setUint32(directoryOffset + 8, (1 << 20) | notes.slidePersistId, true); av.setUint32(directoryOffset + 12, stream.length, true);
    av.setUint32(directoryOffset + 16, (1 << 20) | notes.persistId, true); av.setUint32(directoryOffset + 20, wrapper + 8, true);
    appended.set(oldEditBytes, newEditOffset);
    av.setUint32(newEditOffset + 16, oldEditOffset, true); av.setUint32(newEditOffset + 20, directoryOffset, true);
    const nextUser = user.slice(); viewOf(nextUser).setUint32(16, newEditOffset, true);
    const withNewUser = replaceCompoundFileStream(input, ['Current User'], nextUser);
    const resized = resizeCompoundFileStream(withNewUser, ['PowerPoint Document'], appended);
    expect(resized.ok).toBe(true); if (!resized.ok) return;
    const d = new PptDocument(resized.bytes), t = body(d);
    expect(t.text).toBe('Synthetic notes retained.'); t.text = t.text; expect(d.dirty).toBe(false);
    expect(() => { t.text = 'Synthetic notes modified.'; }).toThrow(/overlaps another live persist object/);
    expect(d.serialize()).toEqual(resized.bytes); expect(d.revision).toBe(0);
  });
  it('refuses a notes container extended over otherwise valid save-directory/user-edit metadata', () => {
    const input = mutate((v, notes) => {
      const user = readCompoundFileStream(load(), ['Current User'])!, edit = readRecordOrThrow(v, viewOf(user).getUint32(16, true));
      const note = readRecordOrThrow(v, notes[1]!.headerOffset);
      v.setUint32(note.headerOffset + 4, edit.dataOffset + edit.recLen - note.dataOffset, true);
    });
    const d = new PptDocument(input), t = body(d, 1);
    expect(t.text).toBe('Unicode notes retained.'); t.text = t.text; expect(d.dirty).toBe(false);
    expect(() => { t.text = 'Unicode notes modified.'; }).toThrow(/save-history metadata/);
    expect(d.serialize()).toEqual(input); expect(d.revision).toBe(0);
  });
  it('rejects malformed notes atom header and odd UTF16 text records', () => {
    const wrongHeader = mutate((v, notes) => v.setUint16(notes[0]!.headerOffset + 8, 0, true));
    expect(() => readPptSlideNotes(wrongHeader)).toThrow(/NotesAtom/);
    const oddWide = mutate((v, notes) => v.setUint16(notes[0]!.texts[0]!.headerOffset + 2, RT.TextCharsAtom, true));
    expect(() => readPptSlideNotes(oddWide)).toThrow(/text atom/);
  });
  it.each([{ maxRecords: 10 }, { maxNotes: 1 }, { maxTextBytes: 1 }, { maxRecords: 0 }, { maxNotes: Infinity }])('enforces resource bounds %j', limit => {
    expect(() => readPptSlideNotes(load(), limit)).toThrow(/budget|limit/);
  });
  it('represents absent notes as undefined without inventing an empty editable page', () => {
    const classic = new Uint8Array(readFileSync(new URL('./fixtures/ppt/classic-rectangle.ppt', import.meta.url)));
    const slide = new PptDocument(classic).slides[0]!;
    expect(slide.notes).toBeUndefined(); expect(slide.notesStatus).toBe('absent'); expect(slide.notesDiagnostic).toBeUndefined();
  });
});
