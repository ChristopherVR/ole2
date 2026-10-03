/** [MS-DOC] 2.2.5, 2.6.1: direct CHPX exceptions, not resolved style formatting.
 * Unknown SPRMs remain opaque. Editing forces an absolute value in an exclusive
 * existing bold/italic or font-size operand, without rebuilding formatting pages. */
import { unwrapDocBytes } from './ole-document-doc-cfb.js';
import { readDocFib } from './ole-document-doc-fib.js';
import { parseBteTable, extractPapxBlob } from './ole-document-doc-fkp.js';
import { parseDocClx, decodePiecesText } from './ole-document-doc-pieces.js';
import { DEFAULT_OLE_DOC_PROCESSING_LIMITS } from './ole-document-doc-editor.js';

export interface DocSprm {
	readonly opcode: number;
	readonly operand: readonly number[];
}
export interface ParsedDocCharacterRun {
	readonly cpStart: number;
	readonly cpEnd: number;
	readonly text: string;
	readonly directBold: boolean | undefined;
	readonly directItalic: boolean | undefined;
	readonly directFontSizePoints: number | undefined;
	readonly directUnderline: 'none' | 'single' | 'double' | undefined;
	readonly sprms: readonly DocSprm[];
}
type Prl = { opcode: number; operand: number[]; at: number };
type PhysicalRun = { start: number; end: number; blob: number; blobEnd: number; prls: Prl[] };
type MappedRun = ParsedDocCharacterRun & { physical: PhysicalRun; prm: number };
const LIMIT = 65_536;
const SAFE_OPERANDS = new Set([
	0x0835, 0x0836, 0x0837, 0x0838, 0x0839, 0x083a, 0x083b, 0x083c,
	0x2a3e, 0x2a42, 0x4a43, 0x4845, 0x2a48, 0x2a53,
	0x4a4f, 0x4a50, 0x4a51, 0x486d, 0x486e, 0x286f,
	0x6870, 0x6815, 0x6816, 0x6817,
]);

/** Variable-size special paragraph/table operands are deliberately refused. */
function parsePrls(word: Uint8Array, start: number, end: number, budget: { remaining: number }): Prl[] {
	const view = new DataView(word.buffer, word.byteOffset, word.byteLength), out: Prl[] = [];
	for (let at = start; at < end;) {
		if (budget.remaining <= 0) throw new Error('Formatting SPRM limit');
		budget.remaining--;
		if (at + 2 > end) throw new Error('Truncated SPRM');
		const opcode = view.getUint16(at, true), spra = opcode >>> 13;
		let count = [1, 1, 2, 4, 2, 2, 0, 3][spra]!;
		if (spra === 6) {
			if (opcode === 0xd608 || opcode === 0xc615 || at + 3 > end) throw new Error('Unsupported SPRM operand');
			count = 1 + word[at + 2]!;
		}
		if (at + 2 + count > end) throw new Error('Truncated SPRM operand');
		out.push({ opcode, operand: Array.from(word.subarray(at + 2, at + 2 + count)), at: at + 2 });
		at += 2 + count;
	}
	return out;
}

function parsePhysicalRuns(word: Uint8Array, table: ReturnType<typeof parseBteTable>): PhysicalRun[] {
	if (table.pns.length > LIMIT) throw new Error('Formatting page limit');
	const view = new DataView(word.buffer, word.byteOffset, word.byteLength), out: PhysicalRun[] = [], budget = { remaining: 262_144 };
	for (let bte = 0; bte < table.pns.length; bte++) {
		const page = (table.pns[bte]! & 0x3fffff) * 512;
		if (page > word.length - 512) throw new Error('Truncated CHPX page');
		const count = word[page + 511]!, header = (count + 1) * 4 + count;
		if (!count || header > 511 || out.length + count > LIMIT) throw new Error('Formatting run limit');
		for (let i = 0; i < count; i++) {
			const start = view.getInt32(page + i * 4, true), end = view.getInt32(page + (i + 1) * 4, true);
			if (start < 0 || end <= start || start < table.fcs[bte]! || end > table.fcs[bte + 1]!)
				throw new Error('Invalid CHPX FC range');
			const offset = word[page + (count + 1) * 4 + i]! * 2;
			if (offset && (offset < header || offset + 1 + word[page + offset]! > 511)) throw new Error('Invalid CHPX blob');
			const blob = offset ? page + offset : 0, blobEnd = blob ? blob + 1 + word[blob]! : 0;
			out.push({ start, end, blob, blobEnd, prls: blob ? parsePrls(word, blob + 1, blobEnd, budget) : [] });
		}
	}
	out.sort((a, b) => a.start - b.start);
	for (let i = 1; i < out.length; i++) if (out[i]!.start < out[i - 1]!.end) throw new Error('Overlapping CHPX ranges');
	return out;
}

function inspect(input: Uint8Array) {
	const doc = unwrapDocBytes(input);
	if (!doc) throw new Error('Unsupported DOC');
	const word = doc.wordDocBytes, fib = readDocFib(word), limits = DEFAULT_OLE_DOC_PROCESSING_LIMITS;
	if (fib.ccpText > limits.maxCharacters) throw new Error('Character limit');
	const pieces = parseDocClx(doc.tableBytes, fib.clx, limits.maxPieces).pieces;
	const total = pieces.at(-1)?.cpEnd ?? 0;
	if (total < fib.ccpText || total > fib.ccpText + fib.ccpOtherStories + 1) throw new Error('Invalid story lengths');
	const text = decodePiecesText(word, pieces, fib.ccpText);
	// Preflight both descriptor counts before parseBteTable allocates its arrays.
	if (fib.plcfbteChpx.lcb > LIMIT * 8 + 4 || fib.plcfbtePapx.lcb > LIMIT * 8 + 4)
		throw new Error('Formatting table limit');
	const chpx = parseBteTable(doc.tableBytes, fib.plcfbteChpx), papx = parseBteTable(doc.tableBytes, fib.plcfbtePapx);
	papx.pns = papx.pns.map((pn) => pn & 0x3fffff);
	const physical = parsePhysicalRuns(word, chpx), runs: MappedRun[] = [];
	for (const piece of pieces) {
		if (piece.cpStart >= fib.ccpText) break;
		const unit = piece.compressed ? 1 : 2, stop = Math.min(piece.cpEnd, fib.ccpText);
		let fc = piece.fc, cp = piece.cpStart;
		let lo = 0, hi = physical.length;
		while (lo < hi) { const mid = (lo + hi) >>> 1; if (physical[mid]!.end <= fc) lo = mid + 1; else hi = mid; }
		while (cp < stop) {
			const row = physical[lo++];
			if (!row || fc < row.start || fc >= row.end || (row.end - fc) % unit) throw new Error('Uncovered or unaligned text formatting');
			const end = Math.min(stop, cp + (row.end - fc) / unit);
			if (runs.length >= LIMIT) throw new Error('Mapped run limit');
			const direct = (opcode: number) => {
				if (piece.prm !== 0 || row.prls.some((p) => !SAFE_OPERANDS.has(p.opcode))) return undefined;
				const values = row.prls.filter((p) => p.opcode === opcode);
				return values.length === 1 && values[0]!.operand[0]! <= 1 ? values[0]!.operand[0] === 1 : undefined;
			};
			const underline = row.prls.filter((p) => p.opcode === 0x2a3e);
			const kul = underline.length === 1 ? underline[0]!.operand[0] : undefined;
			const size = row.prls.filter((p) => p.opcode === 0x4a43);
			const halfPoints = size.length === 1 ? size[0]!.operand[0]! + size[0]!.operand[1]! * 256 : undefined;
			runs.push({ cpStart: cp, cpEnd: end, text: text.slice(cp, end), directBold: direct(0x0835), directItalic: direct(0x0836),
				directUnderline: piece.prm === 0 && row.prls.every((p) => SAFE_OPERANDS.has(p.opcode)) ? kul === 0 ? 'none' : kul === 1 ? 'single' : kul === 3 ? 'double' : undefined : undefined,
				directFontSizePoints: piece.prm === 0 && halfPoints !== undefined && halfPoints >= 2 && halfPoints <= 3276 && row.prls.every((p) => SAFE_OPERANDS.has(p.opcode)) ? halfPoints / 2 : undefined,
				sprms: Object.freeze(row.prls.map((p) => Object.freeze({ opcode: p.opcode, operand: Object.freeze(p.operand.slice()) }))), physical: row, prm: piece.prm });
			fc += (end - cp) * unit; cp = end;
		}
	}
	return { doc, word, fib, pieces, papx, physical, runs, text };
}

/** Direct exceptions only; absence means inherited/unknown, never false. */
export function readDocCharacterRuns(input: Uint8Array): readonly ParsedDocCharacterRun[] {
	return Object.freeze(inspect(input).runs.map(({ physical: _physical, prm: _prm, ...run }) => Object.freeze(run)));
}

/** Style identifier on each main-story paragraph mark; no style resolution. */
export function readDocParagraphStyleIndices(input: Uint8Array): readonly (number | undefined)[] {
	const parsed = inspect(input), styles: (number | undefined)[] = [];
	let pieceIndex = 0;
	for (let cp = 0; cp < parsed.text.length; cp++) {
		if (parsed.text[cp] !== '\r') continue;
		while (parsed.pieces[pieceIndex]!.cpEnd <= cp) pieceIndex++;
		if (styles.length >= LIMIT) throw new Error('Paragraph formatting limit');
		const piece = parsed.pieces[pieceIndex]!;
		const fc = piece.fc + (cp - piece.cpStart) * (piece.compressed ? 1 : 2);
		const blob = extractPapxBlob(parsed.word, parsed.papx, fc), at = blob[0] === 0 ? 2 : 1;
		styles.push(piece.prm === 0 && blob.length >= at + 2 ? blob[at]! + blob[at + 1]! * 256 : undefined);
	}
	return Object.freeze(styles);
}

/** Refuses shared runs/blobs and opaque semantics. Style-dependent operands
 * are replaced with an absolute value, without inferring their previous state. */
export function writeDocCharacterRunFlag(input: Uint8Array, cpStart: number, cpEnd: number, flag: 'bold' | 'italic', value: boolean): Uint8Array {
	if (typeof value !== 'boolean' || !Number.isSafeInteger(cpStart) || !Number.isSafeInteger(cpEnd) || (flag !== 'bold' && flag !== 'italic')) throw new Error('invalid-formatting');
	return writeExclusiveOperand(input, cpStart, cpEnd, flag === 'bold' ? 0x0835 : 0x0836, Number(value));
}

/** [MS-DOC] 2.6.1 sprmCHps: unsigned two-byte half-points, 2..3276.
 * Only an existing exclusive direct slot is writable; styles remain unresolved. */
export function writeDocCharacterRunFontSize(input: Uint8Array, cpStart: number, cpEnd: number, points: number): Uint8Array {
	if (typeof points !== 'number' || !Number.isFinite(points) || !Number.isInteger(points * 2) || points < 1 || points > 1638 ||
		!Number.isSafeInteger(cpStart) || !Number.isSafeInteger(cpEnd)) throw new Error('invalid-formatting');
	return writeExclusiveOperand(input, cpStart, cpEnd, 0x4a43, points * 2);
}

export function writeDocCharacterRunUnderline(input: Uint8Array, cpStart: number, cpEnd: number, value: 'none' | 'single' | 'double'): Uint8Array {
	if (typeof value !== 'string' || !['none', 'single', 'double'].includes(value) || !Number.isSafeInteger(cpStart) || !Number.isSafeInteger(cpEnd)) throw new Error('invalid-formatting');
	return writeExclusiveOperand(input, cpStart, cpEnd, 0x2a3e, value === 'none' ? 0 : value === 'single' ? 1 : 3);
}

function writeExclusiveOperand(input: Uint8Array, cpStart: number, cpEnd: number, opcode: number, value: number): Uint8Array {
	const parsed = inspect(input), target = parsed.runs.find((run) => run.cpStart === cpStart && run.cpEnd === cpEnd);
	if (!target || target.prm !== 0 || /[\x00-\x1f]/u.test(target.text)) throw new Error('unsupported-formatting');
	const row = target.physical;
	const extension = parsed.fib.fibRgFcLcbOffset + parsed.fib.fibRgFcLcbCount * 8;
	const fibEnd = extension + 2 + new DataView(parsed.word.buffer, parsed.word.byteOffset, parsed.word.byteLength).getUint16(extension, true) * 2;
	if (row.blob < fibEnd) throw new Error('aliased-formatting');
	if (row.prls.some((p) => !SAFE_OPERANDS.has(p.opcode))) throw new Error('opaque-formatting');
	const operands = row.prls.filter((p) => p.opcode === opcode);
	if (operands.length !== 1) throw new Error('inherited-formatting');
	const operand = operands[0]!, size = opcode === 0x4a43;
	const previous = operand.operand[0]! + (size ? operand.operand[1]! * 256 : 0);
	if (size ? operand.operand.length !== 2 || previous < 2 || previous > 3276 : opcode === 0x2a3e ? ![0, 1, 3].includes(previous) : ![0, 1, 128, 129].includes(previous)) throw new Error('inherited-formatting');
	// A partial piece/paragraph view must not mutate another logical character.
	const piece = parsed.pieces.find((p) => p.cpStart <= cpStart && p.cpEnd >= cpEnd)!;
	const unit = piece.compressed ? 1 : 2, start = piece.fc + (cpStart - piece.cpStart) * unit;
	if (start !== row.start || start + (cpEnd - cpStart) * unit !== row.end ||
		parsed.runs.filter((r) => r.physical === row).length !== 1 ||
		parsed.pieces.some((p) => p !== piece && p.fc < row.end && row.start < p.fc + (p.cpEnd - p.cpStart) * (p.compressed ? 1 : 2))) throw new Error('shared-formatting');
	if (parsed.physical.some((r) => r !== row && r.blob && r.blob < row.blobEnd && row.blob < r.blobEnd) ||
		parsed.papx.pns.includes(Math.floor(row.blob / 512)) ||
		parsed.pieces.some((p) => p.fc < row.blobEnd && row.blob < p.fc + (p.cpEnd - p.cpStart) * (p.compressed ? 1 : 2))) throw new Error('aliased-formatting');
	// Section exceptions are also stored in WordDocument; reject known aliases.
	const sed = parsed.fib.sed;
	if (sed.lcb > LIMIT * 16 + 4) throw new Error('resource-limit');
	if (sed.lcb) {
		if (sed.lcb < 4 || (sed.lcb - 4) % 16 || sed.fc > parsed.doc.tableBytes.length - sed.lcb) throw new Error('invalid-formatting');
		const n = (sed.lcb - 4) / 16, tableView = new DataView(parsed.doc.tableBytes.buffer, parsed.doc.tableBytes.byteOffset, parsed.doc.tableBytes.byteLength);
		const wordView = new DataView(parsed.word.buffer, parsed.word.byteOffset, parsed.word.byteLength);
		for (let i = 0; i < n; i++) {
			const at = tableView.getInt32(sed.fc + (n + 1) * 4 + i * 12 + 2, true);
			if (at === -1) continue;
			if (at < 0 || at > parsed.word.length - 2) throw new Error('invalid-formatting');
			const end = at + 2 + wordView.getUint16(at, true);
			if (end > parsed.word.length) throw new Error('invalid-formatting');
			if (at < row.blobEnd && row.blob < end) throw new Error('aliased-formatting');
		}
	}
	if (previous === value) return input;
	const word = parsed.word.slice(); word[operand.at] = value & 255;
	if (size) word[operand.at + 1] = value >>> 8;
	return parsed.doc.rewrap(word, parsed.doc.tableBytes);
}
