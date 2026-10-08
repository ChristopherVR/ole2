/** Active-slide text inspection and preservation-first, fixed-slot edits for MS-PPT. */
import { readCompoundFileStream, replaceCompoundFileStream } from './ole2-stream-edit.js';
import { PptParseError, readRecordOrThrow, type PptRecord } from './legacy-ppt-record-stream.js';
import { RT, OA, HEADER_TOKEN_PLAIN, HEADER_TOKEN_ENCRYPTED } from './legacy-ppt-record-types.js';
import { buildPersistDirectory, parseUserEditAtom } from './ppt/persist-directory.js';

export class PptTextError extends PptParseError {
	constructor(public readonly code: 'corrupt' | 'encrypted' | 'unsupported-version', message: string) {
		super(message); this.name = 'PptTextError';
	}
}
export interface PptTextAtom {
	text: string;
	/** Byte offset in the PowerPoint Document stream; diagnostic, not an edit identity. */
	headerOffset: number;
	encoding: 'utf16' | 'compressed-unicode';
}
export interface PptSlideText {
	slideId: number;
	persistId: number;
	/** Outline text followed by inline shape text; no masters, notes or stale saves. */
	texts: PptTextAtom[];
}
export interface PptTextDocument {
	slides: PptSlideText[];
	/** These elements are preserved, but are not decoded by this text API. */
	unsupported: readonly string[];
}

export interface PptTextReadLimits { maxRecords?: number; maxTextBytes?: number; maxSlides?: number }
const DEFAULT_LIMITS = { maxRecords: 100000, maxTextBytes: 16 * 1024 * 1024, maxSlides: 10000 };
interface TextBudget { records: number; textBytes: number; limits: typeof DEFAULT_LIMITS }

const viewOf = (bytes: Uint8Array) => new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
function recordAt(view: DataView, offset: number, type?: number): PptRecord {
	const rec = readRecordOrThrow(view, offset);
	if (rec.dataOffset + rec.recLen > view.byteLength || (type !== undefined && rec.recType !== type))
		throw new PptTextError('corrupt', `Invalid record at ${offset}`);
	return rec;
}
function children(view: DataView, parent: PptRecord, budget: TextBudget): PptRecord[] {
	const result: PptRecord[] = [];
	const end = parent.dataOffset + parent.recLen;
	for (let pos = parent.dataOffset; pos < end;) {
		if (end - pos < 8) throw new PptTextError('corrupt', 'Truncated child header');
		if (++budget.records > budget.limits.maxRecords) throw new PptTextError('corrupt', 'PPT text resource limit exceeded: records');
		const rec = recordAt(view, pos);
		if (rec.dataOffset + rec.recLen > end) throw new PptTextError('corrupt', 'Child exceeds parent');
		result.push(rec); pos = rec.dataOffset + rec.recLen;
	}
	return result;
}
function textAtom(view: DataView, rec: PptRecord, budget: TextBudget): PptTextAtom | undefined {
	if (rec.recType !== RT.TextCharsAtom && rec.recType !== RT.TextBytesAtom) return undefined;
	if (rec.recVer !== 0 || rec.recInstance !== 0 || (rec.recType === RT.TextCharsAtom && rec.recLen % 2))
		throw new PptTextError('corrupt', 'Invalid text atom');
	budget.textBytes += rec.recLen;
	if (budget.textBytes > budget.limits.maxTextBytes) throw new PptTextError('corrupt', 'PPT text resource limit exceeded: text bytes');
	const wide = rec.recType === RT.TextCharsAtom;
	let text = '';
	for (let pos = rec.dataOffset; pos < rec.dataOffset + rec.recLen; pos += wide ? 2 : 1) {
		const value = wide ? view.getUint16(pos, true) : view.getUint8(pos);
		if (value === 0) throw new PptTextError('corrupt', 'NUL in text atom');
		// TextBytesAtom contains low UTF-16 bytes, NOT Windows-1252 ([MS-PPT] 2.9.43).
		text += String.fromCharCode(value);
	}
	return { text, headerOffset: rec.headerOffset, encoding: wide ? 'utf16' : 'compressed-unicode' };
}
function inlineText(view: DataView, parent: PptRecord, budget: TextBudget, ambiguousAtoms: Set<number>): { texts: PptTextAtom[]; mirrors: PptRecord[] } {
	const result: PptTextAtom[] = [];
	const mirrors: PptRecord[] = [];
	const stack: Array<{ parent: PptRecord; depth: number; shape?: PptRecord }> = [{ parent, depth: 0 }];
	while (stack.length) {
		const current = stack.pop()!;
		if (current.depth > 128) throw new PptTextError('corrupt', 'Record nesting exceeds limit');
		const records = children(view, current.parent, budget);
		markAmbiguousTextGroups(records, current.parent, ambiguousAtoms);
		for (let i = records.length - 1; i >= 0; i--) {
			const rec = records[i]!;
			// ClientTextbox has atom framing but contains PPT records ([MS-ODRAW]).
			if (rec.recVer === 15 || rec.recType === OA.ClientTextbox)
				stack.push({ parent: rec, depth: current.depth + 1, shape: rec.recType === OA.SpContainer ? rec : current.shape });
		}
		for (const rec of records) {
			const atom = textAtom(view, rec, budget); if (atom) result.push(atom);
			if (rec.recType === OA.FOPT || rec.recType === OA.TertiaryFOPT) {
				if (rec.recInstance * 6 > rec.recLen) throw new PptTextError('corrupt', 'Invalid drawing property table');
				for (let i = 0; i < rec.recInstance; i++) {
					// metroBlob can override binary fallback content in modern PowerPoint.
					if ((view.getUint16(rec.dataOffset + i * 6, true) & 0x3fff) === 0x3a9) mirrors.push(current.shape ?? parent);
				}
			}
		}
	}
	return { texts: result.sort((a, b) => a.headerOffset - b.headerOffset), mirrors };
}
/** A TextHeader owns one text atom; malformed groups remain readable but not editable. */
function markAmbiguousTextGroups(records: PptRecord[], parent: PptRecord, ambiguous: Set<number>): void {
	const textParent = parent.recType === OA.ClientTextbox || (parent.recType === RT.SlideListWithText && parent.recInstance === 0);
	let atoms: number[] = [], validHeader = false;
	const finish = () => { if (!validHeader || atoms.length !== 1) for (const offset of atoms) ambiguous.add(offset); atoms = []; };
	for (const rec of records) {
		if (rec.recType === RT.TextHeaderAtom || rec.recType === RT.SlidePersistAtom) {
			finish();
			validHeader = textParent && rec.recType === RT.TextHeaderAtom && rec.recVer === 0 && rec.recLen === 4
				&& rec.recInstance <= (parent.recType === OA.ClientTextbox ? 0 : 5);
		} else if (rec.recType === RT.TextCharsAtom || rec.recType === RT.TextBytesAtom) atoms.push(rec.headerOffset);
	}
	finish();
}
function inspect(input: Uint8Array, limits: PptTextReadLimits = {}): { model: PptTextDocument; stream: Uint8Array; mirroredAtoms: Set<number>; ambiguousAtoms: Set<number> } {
	const checkedLimits = { ...DEFAULT_LIMITS, ...limits };
	if (!Object.values(checkedLimits).every(value => Number.isSafeInteger(value) && value > 0)) throw new PptTextError('corrupt', 'Invalid text resource limit');
	const budget: TextBudget = { records: 0, textBytes: 0, limits: checkedLimits };
	const ambiguousAtoms = new Set<number>();
	const current = readCompoundFileStream(input, ['Current User']);
	const stream = readCompoundFileStream(input, ['PowerPoint Document']);
	if (!current || !stream) throw new PptTextError('corrupt', 'Missing or malformed PowerPoint streams');
	const cv = viewOf(current), dv = viewOf(stream);
	const cu = recordAt(cv, 0, RT.CurrentUserAtom), d = cu.dataOffset;
	if (cu.recVer !== 0 || cu.recInstance !== 0 || cu.recLen < 24 || cv.getUint32(d, true) !== 20)
		throw new PptTextError('corrupt', 'Invalid CurrentUserAtom');
	const nameLength = cv.getUint16(d + 12, true);
	if (nameLength > 255 || 24 + nameLength > cu.recLen)
		throw new PptTextError('corrupt', 'Invalid CurrentUserAtom username bounds');
	const token = cv.getUint32(d + 4, true);
	if (token === HEADER_TOKEN_ENCRYPTED) throw new PptTextError('encrypted', 'Encrypted PPT text is unsupported');
	if (token !== HEADER_TOKEN_PLAIN) throw new PptTextError('corrupt', 'Unknown encryption token');
	if (cv.getUint16(d + 14, true) !== 0x3f4 || cv.getUint8(d + 16) !== 3 || cv.getUint8(d + 17) !== 0)
		throw new PptTextError('unsupported-version', 'Expected PowerPoint 97-2003 storage');
	const { currentEdit, directory } = buildPersistDirectory(dv, cv.getUint32(d + 8, true), { maxEntries: checkedLimits.maxRecords });
	if (currentEdit.encryptSessionPersistIdRef !== undefined)
		throw new PptTextError('encrypted', 'Encrypted persist objects are unsupported');
	const docOffset = directory.get(currentEdit.docPersistIdRef);
	if (docOffset === undefined) throw new PptTextError('corrupt', 'Missing active document');
	const doc = recordAt(dv, docOffset, RT.Document);
	if (doc.recVer !== 15) throw new PptTextError('corrupt', 'Invalid document container');
	const slides: PptSlideText[] = [];
	const slideIds = new Set<number>(), persistIds = new Set<number>();
	for (const list of children(dv, doc, budget)) {
		if (list.recType !== RT.SlideListWithText || list.recInstance !== 0) continue;
		if (list.recVer !== 15) throw new PptTextError('corrupt', 'Invalid slide list');
		let slide: PptSlideText | undefined;
		const records = children(dv, list, budget);
		markAmbiguousTextGroups(records, list, ambiguousAtoms);
		for (const rec of records) {
			if (rec.recType === RT.SlidePersistAtom) {
				if (rec.recVer !== 0 || rec.recInstance !== 0 || rec.recLen !== 20)
					throw new PptTextError('corrupt', 'Invalid slide persist atom');
				const persistId = dv.getUint32(rec.dataOffset, true);
				slide = { persistId, slideId: dv.getUint32(rec.dataOffset + 12, true), texts: [] };
				if (slideIds.has(slide.slideId) || persistIds.has(persistId))
					throw new PptTextError('corrupt', 'Duplicate slide identity');
				slideIds.add(slide.slideId); persistIds.add(persistId);
				if (slides.length >= checkedLimits.maxSlides) throw new PptTextError('corrupt', 'PPT text resource limit exceeded: slides');
				slides.push(slide);
			} else {
				const atom = textAtom(dv, rec, budget);
				if (atom) { if (!slide) throw new PptTextError('corrupt', 'Unowned outline text'); slide.texts.push(atom); }
			}
		}
	}
	const mirroredAtoms = new Set<number>();
	for (const slide of slides) {
		const offset = directory.get(slide.persistId);
		if (offset === undefined) throw new PptTextError('corrupt', 'Missing active slide');
		const rec = recordAt(dv, offset, RT.Slide);
		if (rec.recVer !== 15) throw new PptTextError('corrupt', 'Invalid slide container');
		const inline = inlineText(dv, rec, budget, ambiguousAtoms);
		// Outline-to-shape refs are not resolved, so a mirrored slide's outline is unsafe.
		if (inline.mirrors.length) for (const atom of slide.texts) mirroredAtoms.add(atom.headerOffset);
		slide.texts.push(...inline.texts);
		for (const atom of inline.texts) if (inline.mirrors.some(m => atom.headerOffset >= m.dataOffset && atom.headerOffset < m.dataOffset + m.recLen))
			mirroredAtoms.add(atom.headerOffset);
	}
	return { stream, mirroredAtoms, ambiguousAtoms, model: { slides, unsupported: ['formatting', 'geometry', 'pictures', 'masters', 'notes', 'animations', 'embedded-objects'] } };
}
/** Inspect active slide text without evaluating links, macros or embedded objects. */
export function readPptSlideTexts(input: Uint8Array, limits: PptTextReadLimits = {}): PptTextDocument {
	try { return inspect(input, limits).model; }
	catch (error) {
		if (error instanceof PptTextError) throw error;
		throw new PptTextError('corrupt', error instanceof Error ? error.message : 'Malformed PPT');
	}
}
export type PptTextEditResult =
	| { status: 'edited' | 'unchanged'; bytes: Uint8Array }
	| { status: 'unsupported'; bytes: Uint8Array; reason: string };
/** Guard complete owning containers, including opaque bytes not exposed as text. */
function textOwnershipRefusal(input: Uint8Array, stream: Uint8Array, persistId: number, atom: PptTextAtom): string | undefined {
	const user = readCompoundFileStream(input, ['Current User']);
	if (!user) throw new PptParseError('Missing Current User stream');
	const view = viewOf(stream), uv = viewOf(user), cu = readRecordOrThrow(uv, 0);
	const editOffset = uv.getUint32(cu.dataOffset + 8, true);
	const { directory, currentEdit } = buildPersistDirectory(view, editOffset);
	const slideOffset = directory.get(persistId);
	if (slideOffset === undefined) return 'Missing active slide persist object';
	const slide = readRecordOrThrow(view, slideOffset), text = readRecordOrThrow(view, atom.headerOffset);
	const textEnd = text.dataOffset + text.recLen;
	const owners = [persistId];
	// Outline atoms live in the active Document rather than the Slide container.
	if (atom.headerOffset < slide.dataOffset || textEnd > slide.dataOffset + slide.recLen) {
		const docOffset = directory.get(currentEdit.docPersistIdRef);
		if (docOffset === undefined) return 'Missing outline text owner';
		const doc = readRecordOrThrow(view, docOffset);
		if (atom.headerOffset < doc.dataOffset || textEnd > doc.dataOffset + doc.recLen) return 'Text is outside its active owning containers';
		owners.push(currentEdit.docPersistIdRef);
	}
	for (const ownerId of owners) {
		const owned = readRecordOrThrow(view, directory.get(ownerId)!);
		const start = owned.headerOffset, end = owned.dataOffset + owned.recLen;
		if (end > view.byteLength) return 'Unbounded text owning persist object';
		const overlaps = (offset: number): boolean => {
			const other = readRecordOrThrow(view, offset), otherEnd = other.dataOffset + other.recLen;
			if (otherEnd > view.byteLength) throw new PptParseError('Unbounded live persist object prevents safe text edit');
			return start < otherEnd && other.headerOffset < end;
		};
		for (const [id, offset] of directory) if (id !== ownerId && overlaps(offset)) return 'Text owning container overlaps another live persist object';
		const visited = new Set<number>();
		for (let offset = editOffset; offset;) {
			if (visited.has(offset) || visited.size >= 10000) return 'Invalid text save-history ownership';
			visited.add(offset);
			const edit = parseUserEditAtom(view, offset);
			if (overlaps(offset) || overlaps(edit.offsetPersistDirectory)) return 'Text owning container overlaps save-history metadata';
			offset = edit.offsetLastEdit;
		}
	}
	return undefined;
}
/** Replace an existing active text atom in its fixed character slots. All CFB bytes
 * outside its payload remain unchanged, including unknown records and streams.
 * Styles and hyperlink ranges stay at their original character offsets. */
export function editPptSlideText(input: Uint8Array, edit: {
	slideIndex: number; textIndex: number; expectedText: string; text: string;
}): PptTextEditResult {
	const reject = (reason: string): PptTextEditResult => ({ status: 'unsupported', bytes: input, reason });
	try {
		// Capture caller accessors once so the ownership guard checks the same
		// selected text and replacement that validation and encoding use.
		edit = { slideIndex: edit.slideIndex, textIndex: edit.textIndex, expectedText: edit.expectedText, text: edit.text };
		if (!Number.isSafeInteger(edit.slideIndex) || !Number.isSafeInteger(edit.textIndex) || edit.slideIndex < 0 || edit.textIndex < 0)
			return reject('Invalid text location');
		const { stream, model, mirroredAtoms, ambiguousAtoms } = inspect(input);
		const atom = model.slides[edit.slideIndex]?.texts[edit.textIndex];
		if (!atom || atom.text !== edit.expectedText) return reject('Text location or expected text does not match');
		if (edit.text === atom.text) return { status: 'unchanged', bytes: input };
		if (model.slides.reduce((count, slide) => count + slide.texts.filter(text => text.headerOffset === atom.headerOffset).length, 0) !== 1)
			return reject('Text atom is shared by multiple active slide positions');
		if (ambiguousAtoms.has(atom.headerOffset)) return reject('Text atom has no unique valid TextHeader owner');
		if (mirroredAtoms.has(atom.headerOffset))
			return reject('Text has an OOXML mirror that could override the binary text');
		if (edit.text.length !== atom.text.length) return reject('Replacement must keep the UTF-16 character count');
		for (let i = 0; i < edit.text.length; i++) {
			const old = atom.text.charCodeAt(i), value = edit.text.charCodeAt(i);
			if (value === 0 || ((old < 32 || value < 32 || old === 42 || value === 42) && old !== value))
				return reject('Control characters and field markers must stay at the same offsets');
			if (atom.encoding === 'compressed-unicode' && value > 255) return reject('Character does not fit the existing encoding');
			if (value >= 0xd800 && value <= 0xdbff && !(edit.text.charCodeAt(i + 1) >= 0xdc00 && edit.text.charCodeAt(i + 1) <= 0xdfff))
				return reject('Unpaired UTF-16 surrogate');
			if (value >= 0xdc00 && value <= 0xdfff && !(edit.text.charCodeAt(i - 1) >= 0xd800 && edit.text.charCodeAt(i - 1) <= 0xdbff))
				return reject('Unpaired UTF-16 surrogate');
		}
		const ownership = textOwnershipRefusal(input, stream, model.slides[edit.slideIndex]!.persistId, atom);
		if (ownership) return reject(ownership);
		const output = stream.slice(), view = viewOf(output), start = atom.headerOffset + 8;
		for (let i = 0; i < edit.text.length; i++) {
			if (atom.encoding === 'utf16') view.setUint16(start + i * 2, edit.text.charCodeAt(i), true);
			else output[start + i] = edit.text.charCodeAt(i);
		}
		const bytes = replaceCompoundFileStream(input, ['PowerPoint Document'], output);
		return bytes === input ? reject('Compound stream cannot be safely replaced') : { status: 'edited', bytes };
	} catch (error) { return reject(error instanceof Error ? error.message : 'Malformed PPT'); }
}
