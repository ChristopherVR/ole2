/** Bounded existing StyleTextPropAtom character runs ([MS-PPT] 2.9.14,
 * 2.9.19 and 2.9.46-48). Unknown formatting remains opaque. */
import { readCompoundFileStream, replaceCompoundFileStream } from './ole2-stream-edit.js';
import { readPptSlideTexts, type PptTextEditResult, type PptTextAtom } from './legacy-ppt-text.js';
import { readPptSlideShapes } from './legacy-ppt-shape-reader.js';
import { readRecordOrThrow, PptParseError, type PptRecord } from './legacy-ppt-record-stream.js';
import { RT, OA } from './legacy-ppt-record-types.js';
import { buildPersistDirectory, parseUserEditAtom } from './ppt/persist-directory.js';

export interface PptCharacterRunRecord {
	start: number; end: number; rawCount: number; mask: number;
	styleHeaderOffset: number; runOffset: number; fontSizeOffset?: number;
	directFontSizePoints?: number;
}
export interface PptTextRunRecord {
	slideId: number; persistId: number; headerOffset: number; text: string;
	status: 'decoded' | 'absent' | 'unsupported'; diagnostic?: string;
	runs: PptCharacterRunRecord[];
}
const viewOf = (bytes: Uint8Array) => new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
const MAX_RECORDS = 100000, MAX_RUNS = 20000, MAX_STYLE_BYTES = 16 * 1024 * 1024;
export interface PptCharacterRunReadLimits { maxRecords?: number; maxRuns?: number; maxStyleBytes?: number }
function children(view: DataView, parent: PptRecord): PptRecord[] {
	const result: PptRecord[] = [], end = parent.dataOffset + parent.recLen;
	if (end > view.byteLength) throw new Error('Unbounded formatting parent');
	for (let offset = parent.dataOffset; offset < end;) {
		if (result.length >= MAX_RECORDS || end - offset < 8) throw new Error('Formatting record bounds exceeded');
		const record = readRecordOrThrow(view, offset);
		if (record.dataOffset + record.recLen > end) throw new Error('Formatting child exceeds its parent');
		result.push(record); offset = record.dataOffset + record.recLen;
	}
	return result;
}
function decodeStyle(view: DataView, style: PptRecord, textLength: number, budget: { runs: number; maxRuns: number }): PptCharacterRunRecord[] {
	if (style.recVer !== 0 || style.recInstance !== 0) throw new Error('Invalid StyleTextPropAtom header');
	let pos = style.dataOffset;
	const end = pos + style.recLen, total = textLength + 1;
	const take = (bytes: number): number => { if (bytes > end - pos) throw new Error('Truncated formatting exception'); const offset = pos; pos += bytes; return offset; };
	const u16 = () => view.getUint16(take(2), true), u32 = () => view.getUint32(take(4), true);
	const count = (sum: number): number => {
		if (++budget.runs > budget.maxRuns) throw new Error('Character/paragraph run budget exceeded');
		const value = u32(); if (value === 0 || value > total - sum) throw new Error('Invalid formatting run count'); return value;
	};
	// The implicit final paragraph mark has a character slot in native PPT.
	for (let sum = 0; sum < total;) {
		sum += count(sum); if (u16() > 4) throw new Error('Invalid paragraph indentation');
		const mask = u32();
		// StyleTextPropAtom forbids ruler fields and v9 extension fields.
		if (mask & ~0x002f7aff) throw new Error('Unsupported paragraph formatting mask');
		if (mask & 0xf) take(2);
		for (const bit of [0x80, 0x10, 0x40]) if (mask & bit) take(2);
		if (mask & 0x20) take(4);
		for (const bit of [0x800, 0x1000, 0x2000, 0x4000, 0x10000]) if (mask & bit) take(2);
		if (mask & 0xe0000) take(2);
		if (mask & 0x200000) take(2);
	}
	const result: PptCharacterRunRecord[] = [];
	for (let sum = 0; sum < total;) {
		const runOffset = pos, rawCount = count(sum), start = sum; sum += rawCount;
		const mask = u32();
		if (mask & ~0x00efffff) throw new Error('Unsupported character formatting mask');
		if (mask & 0x3eb7) take(2);
		for (const bit of [0x10000, 0x200000, 0x400000, 0x800000]) if (mask & bit) take(2);
		let fontSizeOffset: number | undefined, directFontSizePoints: number | undefined;
		if (mask & 0x20000) {
			fontSizeOffset = take(2); directFontSizePoints = view.getInt16(fontSizeOffset, true);
			if (directFontSizePoints < 1 || directFontSizePoints > 4000) throw new Error('Invalid literal font size');
		}
		if (mask & 0x40000) take(4);
		if (mask & 0x80000) { const position = view.getInt16(take(2), true); if (position < -100 || position > 100) throw new Error('Invalid baseline position'); }
		result.push({ start: Math.min(start, textLength), end: Math.min(sum, textLength), rawCount, mask,
			styleHeaderOffset: style.headerOffset, runOffset, fontSizeOffset, directFontSizePoints });
	}
	if (pos !== end) throw new Error('Unconsumed StyleTextPropAtom bytes');
	return result;
}
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
/** Direct exceptions only; undefined font size means inheritance is unresolved. */
export function readPptCharacterRuns(input: Uint8Array, limits: PptCharacterRunReadLimits = {}): PptTextRunRecord[] {
	const maxRecords = limits.maxRecords ?? MAX_RECORDS, maxRuns = limits.maxRuns ?? MAX_RUNS, maxStyleBytes = limits.maxStyleBytes ?? MAX_STYLE_BYTES;
	if (![maxRecords, maxRuns, maxStyleBytes].every(value => Number.isSafeInteger(value) && value > 0)) throw new Error('Invalid character-format resource limit');
	// Existing shape preflight bounds records and text before allocating strings.
	readPptSlideShapes(input, { maxRecords });
	const texts = readPptSlideTexts(input), stream = readCompoundFileStream(input, ['PowerPoint Document'])!, user = readCompoundFileStream(input, ['Current User'])!;
	const view = viewOf(stream), uv = viewOf(user), cu = readRecordOrThrow(uv, 0);
	const { directory, currentEdit } = buildPersistDirectory(view, uv.getUint32(cu.dataOffset + 8, true));
	const associations = new Map<number, { styles: PptRecord[]; textCount: number }>();
	let records = 0, styleBytes = 0;
	const scan = (root: PptRecord) => {
		const stack = [{ record: root, depth: 0 }];
		while (stack.length) {
			const { record, depth } = stack.pop()!; if (depth > 128) throw new Error('Formatting nesting exceeds limit');
			const list = children(view, record); records += list.length;
			if (records > maxRecords) throw new Error('Formatting record budget exceeded');
			let group: { atoms: number[]; styles: PptRecord[] } | undefined;
			const textParent = record.recType === OA.ClientTextbox || (record.recType === RT.SlideListWithText && record.recInstance === 0);
			const finish = () => { if (group) for (const atom of group.atoms) {
				if (associations.has(atom)) throw new Error('Ambiguous formatting parent');
				associations.set(atom, { styles: group.styles, textCount: group.atoms.length });
			} };
			for (const child of list) {
				if (textParent && (child.recType === RT.TextHeaderAtom || child.recType === RT.SlidePersistAtom)) {
					finish(); group = child.recType === RT.TextHeaderAtom ? { atoms: [], styles: [] } : undefined;
					if (group && (child.recVer !== 0 || child.recInstance > (record.recType === OA.ClientTextbox ? 0 : 5) || child.recLen !== 4)) throw new Error('Invalid formatting TextHeaderAtom');
				}
				else if (group && (child.recType === RT.TextCharsAtom || child.recType === RT.TextBytesAtom)) group.atoms.push(child.headerOffset);
				else if (group && child.recType === RT.StyleTextPropAtom) { styleBytes += child.recLen; if (styleBytes > maxStyleBytes) throw new Error('Formatting payload budget exceeded'); group.styles.push(child); }
				if (child.recVer === 15 || child.recType === OA.ClientTextbox) stack.push({ record: child, depth: depth + 1 });
			} finish();
		}
	};
	let scanDiagnostic: string | undefined;
	try {
		const doc = readRecordOrThrow(view, directory.get(currentEdit.docPersistIdRef)!);
		for (const list of children(view, doc)) if (list.recType === RT.SlideListWithText && list.recInstance === 0) scan(list);
		const visited = new Set<number>();
		for (const slide of texts.slides) { const offset = directory.get(slide.persistId)!; if (!visited.has(offset)) { visited.add(offset); scan(readRecordOrThrow(view, offset)); } }
	} catch (error) { scanDiagnostic = error instanceof Error ? error.message : 'Unsupported formatting graph'; }
	const budget = { runs: 0, maxRuns };
	return texts.slides.flatMap(slide => slide.texts.map(atom => {
		const base = { slideId: slide.slideId, persistId: slide.persistId, headerOffset: atom.headerOffset, text: atom.text };
		if (scanDiagnostic) return { ...base, status: 'unsupported' as const, diagnostic: scanDiagnostic, runs: [] };
		const group = associations.get(atom.headerOffset);
		if (!group) return { ...base, status: 'unsupported' as const, diagnostic: 'Text has no unambiguous TextHeader formatting parent', runs: [] };
		if (!group.styles.length) return { ...base, status: 'absent' as const, runs: [] };
		if (group.textCount !== 1 || group.styles.length !== 1) return { ...base, status: 'unsupported' as const, diagnostic: 'Text or style atom association is ambiguous', runs: [] };
		try { return { ...base, status: 'decoded' as const, runs: decodeStyle(view, group.styles[0]!, atom.text.length, budget) }; }
		catch (error) { return { ...base, status: 'unsupported' as const, diagnostic: error instanceof Error ? error.message : 'Unsupported formatting', runs: [] }; }
	}));
}
/** Update exactly one existing two-byte fontSize field; no new style records. */
export function editPptCharacterFontSize(input: Uint8Array, edit: {
	slideId: number; persistId: number; headerOffset: number; runOffset: number;
	styleHeaderOffset: number; start: number; end: number; expectedFontSizePoints: number | undefined; fontSizePoints: number | undefined;
}): PptTextEditResult {
	const reject = (reason: string): PptTextEditResult => ({ status: 'unsupported', bytes: input, reason });
	try {
		const { slideId, persistId, headerOffset, runOffset, styleHeaderOffset, start, end, expectedFontSizePoints, fontSizePoints } = edit;
		const matches = readPptCharacterRuns(input).filter(text => text.slideId === slideId && text.persistId === persistId && text.headerOffset === headerOffset);
		const record = matches.length === 1 ? matches[0] : undefined;
		const runs = record?.status === 'decoded' ? record.runs.filter(run => run.runOffset === runOffset && run.styleHeaderOffset === styleHeaderOffset && run.start === start && run.end === end) : [];
		const run = runs.length === 1 ? runs[0] : undefined;
		if (!run || run.directFontSizePoints !== expectedFontSizePoints) return reject(record?.diagnostic ?? 'Character run identity or expected font size does not match');
		if (fontSizePoints === run.directFontSizePoints) return { status: 'unchanged', bytes: input };
		if (run.start === run.end) return reject('Run covers only the implicit paragraph mark');
		if (run.fontSizeOffset === undefined) return reject('Font size is inherited; no existing literal slot');
		if (!Number.isSafeInteger(fontSizePoints) || fontSizePoints! < 1 || fontSizePoints! > 4000) return reject('Literal font size must be an integer from 1 to 4000 points');
		const textModel = readPptSlideTexts(input), slides = textModel.slides.filter(slide => slide.slideId === slideId && slide.persistId === persistId);
		const indexes = slides.length === 1 ? slides[0]!.texts.flatMap((atom, index) => atom.headerOffset === headerOffset ? [index] : []) : [];
		if (indexes.length !== 1 || textModel.slides.reduce((count, slide) => count + slide.texts.filter(atom => atom.headerOffset === headerOffset).length, 0) !== 1) return reject('Text atom is shared by multiple active slide positions');
		const shapeSlides = readPptSlideShapes(input).filter(slide => slide.slideId === slideId && slide.persistId === persistId);
		const shapes = shapeSlides.length === 1 ? shapeSlides[0]!.shapes.filter(shape => shape.textReference === 'inline' && shape.textIndexes.includes(indexes[0]!)) : [];
		if (shapes.length !== 1 || shapes[0]!.kind !== 'text-box') return reject('Font-size edits require uniquely linked inline text-box text');
		if (shapes[0]!.geometryRefusal) return reject(`Unsupported text-box formatting context: ${shapes[0]!.geometryRefusal}`);
		const originalStream = readCompoundFileStream(input, ['PowerPoint Document'])!;
		const refusal = textOwnershipRefusal(input, originalStream, persistId, slides[0]!.texts[indexes[0]!]!); if (refusal) return reject(refusal);
		const stream = readCompoundFileStream(input, ['PowerPoint Document'])!.slice(); viewOf(stream).setInt16(run.fontSizeOffset, fontSizePoints!, true);
		const bytes = replaceCompoundFileStream(input, ['PowerPoint Document'], stream); if (bytes === input) return reject('Compound stream cannot be safely replaced');
		const candidate = readPptCharacterRuns(bytes).filter(text => text.slideId === slideId && text.persistId === persistId && text.headerOffset === headerOffset);
		const next = candidate.length === 1 && candidate[0]!.status === 'decoded' ? candidate[0]!.runs.find(item => item.runOffset === runOffset && item.styleHeaderOffset === styleHeaderOffset && item.start === start && item.end === end) : undefined;
		if (!next || next.directFontSizePoints !== fontSizePoints || candidate[0]!.text !== record!.text) return reject('Edited character run failed candidate validation');
		return { status: 'edited', bytes };
	} catch (error) { return reject(error instanceof Error ? error.message : 'Malformed PPT formatting'); }
}

