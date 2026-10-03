/** Bounded direct PAPX exceptions ([MS-DOC] 2.6.2, 2.9.175).
 * Logical alignment does not resolve paragraph styles or physical bidi layout. */
import { unwrapDocBytes } from './ole-document-doc-cfb.js';
import { readDocFib } from './ole-document-doc-fib.js';
import { parseBteTable } from './ole-document-doc-fkp.js';
import { parseDocClx, decodePiecesText } from './ole-document-doc-pieces.js';
import { DEFAULT_OLE_DOC_PROCESSING_LIMITS } from './ole-document-doc-editor.js';
import type { DocParagraphAlignment } from './doc-document.js';

const LIMIT = 65_536;
const SPRM_LIMIT = 262_144;
// PnFkpPapx/PnFkpChpx have a 22-bit pn and 10 undefined bits that MUST be ignored.
const PAGE_NUMBER_MASK = 0x003fffff;
const ALIGNMENTS: readonly DocParagraphAlignment[] = ['start', 'center', 'end', 'justify'];
// No reset, style-change, table or huge-PAPX semantics. Legacy physical alignment
// is understood only as a unique preceding matching compatibility mirror.
const SAFE = new Set([0x2403, 0x2461, 0x6467, 0x2405, 0x2406, 0x2407, 0x240c,
	0x6412, 0xa413, 0xa414, 0x245b, 0x245c, 0x845d, 0x845e, 0x8460, 0x2441, 0x246d]);
type Prl = { opcode: number; operand: Uint8Array; at: number };
type Row = { start: number; end: number; blob: number; blobEnd: number; prls: Prl[] };
type Paragraph = { cpStart: number; cpEnd: number; mark: boolean; row: Row | undefined; prm: number; directAlignment: DocParagraphAlignment | undefined };

function alignmentOperands(row: Row) {
	const values = row.prls.filter((p) => p.opcode === 0x2461), mirrors = row.prls.filter((p) => p.opcode === 0x2403);
	if (values.length !== 1 || values[0]!.operand[0]! > 3 || mirrors.length > 1) return undefined;
	const value = values[0]!, mirror = mirrors[0];
	if (mirror && (mirror.at > value.at || mirror.operand[0] !== value.operand[0])) return undefined;
	return { value, mirror };
}

function parsePrls(word: Uint8Array, start: number, end: number, budget: { remaining: number }): Prl[] {
	const view = new DataView(word.buffer, word.byteOffset, word.byteLength), out: Prl[] = [];
	for (let at = start; at < end;) {
		if (at + 2 > end) throw new Error('invalid-formatting');
		const opcode = view.getUint16(at, true), spra = opcode >>> 13;
		let length = [1, 1, 2, 4, 2, 2, 0, 3][spra]!;
		if (spra === 6) {
			if (opcode === 0xd608 || opcode === 0xc615 || at + 3 > end) throw new Error('unsupported-formatting');
			length = 1 + word[at + 2]!;
		}
		if (at + 2 + length > end) throw new Error('invalid-formatting');
		if (budget.remaining === 0) throw new Error('resource-limit');
		budget.remaining--;
		out.push({ opcode, at: at + 2, operand: word.subarray(at + 2, at + 2 + length) });
		at += 2 + length;
	}
	return out;
}

function inspect(input: Uint8Array) {
	const doc = unwrapDocBytes(input);
	if (!doc) throw new Error('unsupported-formatting');
	const word = doc.wordDocBytes, fib = readDocFib(word), limits = DEFAULT_OLE_DOC_PROCESSING_LIMITS;
	if (fib.ccpText > limits.maxCharacters) throw new Error('resource-limit');
	const pieces = parseDocClx(doc.tableBytes, fib.clx, limits.maxPieces).pieces;
	const total = pieces.at(-1)?.cpEnd ?? 0;
	if (total < fib.ccpText || total > fib.ccpText + fib.ccpOtherStories + 1) throw new Error('invalid-formatting');
	const text = decodePiecesText(word, pieces, fib.ccpText);
	if (fib.plcfbtePapx.lcb > LIMIT * 8 + 4 || fib.plcfbteChpx.lcb > LIMIT * 8 + 4) throw new Error('resource-limit');
	const papx = parseBteTable(doc.tableBytes, fib.plcfbtePapx), chpx = parseBteTable(doc.tableBytes, fib.plcfbteChpx);
	const view = new DataView(word.buffer, word.byteOffset, word.byteLength), rows: Row[] = [];
	const sprmBudget = { remaining: SPRM_LIMIT };
	for (let pageIndex = 0; pageIndex < papx.pns.length; pageIndex++) {
		const page = (papx.pns[pageIndex]! & PAGE_NUMBER_MASK) * 512;
		if (page > word.length - 512) throw new Error('invalid-formatting');
		const count = word[page + 511]!, header = (count + 1) * 4 + count * 13;
		if (!count || header > 511 || rows.length + count > LIMIT) throw new Error('resource-limit');
		for (let i = 0; i < count; i++) {
			const start = view.getInt32(page + i * 4, true), end = view.getInt32(page + (i + 1) * 4, true);
			if (start < papx.fcs[pageIndex]! || end > papx.fcs[pageIndex + 1]! || end <= start) throw new Error('invalid-formatting');
			const offset = word[page + (count + 1) * 4 + i * 13]! * 2;
			if (!offset) { rows.push({ start, end, blob: 0, blobEnd: 0, prls: [] }); continue; }
			if (offset < header || offset + 2 > 511) throw new Error('invalid-formatting');
			const blob = page + offset, cb = word[blob]!, prefix = cb === 0 ? 2 : 1;
			const length = cb === 0 ? 2 + word[blob + 1]! * 2 : cb * 2;
			if ((cb === 0 && !word[blob + 1]) || length < prefix + 2 || offset + length > 511) throw new Error('invalid-formatting');
			const blobEnd = blob + length;
			rows.push({ start, end, blob, blobEnd, prls: parsePrls(word, blob + prefix + 2, blobEnd, sprmBudget) });
		}
	}
	rows.sort((a, b) => a.start - b.start);
	for (let i = 1; i < rows.length; i++) if (rows[i]!.start < rows[i - 1]!.end) throw new Error('invalid-formatting');
	const paragraphs: Paragraph[] = [];
	let cpStart = 0, pieceIndex = 0;
	for (let cp = 0; cp < text.length; cp++) {
		if (text[cp] !== '\r' && cp !== text.length - 1) continue;
		if (paragraphs.length >= LIMIT) throw new Error('resource-limit');
		while (pieces[pieceIndex]!.cpEnd <= cp) pieceIndex++;
		const piece = pieces[pieceIndex]!, fc = piece.fc + (cp - piece.cpStart) * (piece.compressed ? 1 : 2);
		let lo = 0, hi = rows.length;
		while (lo < hi) { const middle = (lo + hi) >>> 1; if (rows[middle]!.end <= fc) lo = middle + 1; else hi = middle; }
		const row = rows[lo];
		if (!row || fc < row.start || fc >= row.end) throw new Error('invalid-formatting');
		const operands = alignmentOperands(row);
		const directAlignment = piece.prm === 0 && row.prls.every((p) => SAFE.has(p.opcode)) && operands ? ALIGNMENTS[operands.value.operand[0]!] : undefined;
		paragraphs.push({ cpStart, cpEnd: cp + 1, mark: text[cp] === '\r', row, prm: piece.prm, directAlignment });
		cpStart = cp + 1;
	}
	return { doc, word, fib, pieces, rows, chpx, paragraphs, text };
}

/** Missing or opaque direct values remain undefined; no inherited defaults. */
export function readDocParagraphAlignments(input: Uint8Array): readonly (DocParagraphAlignment | undefined)[] {
	return Object.freeze(inspect(input).paragraphs.map((p) => p.directAlignment));
}

/** Replace a sole existing sprmPJc in an exclusive PAPX paragraph run. A matching
 * legacy mirror is changed together, only for direction-independent center/justify. */
export function writeDocParagraphAlignment(input: Uint8Array, index: number, alignment: DocParagraphAlignment): Uint8Array {
	if (!Number.isSafeInteger(index) || index < 0 || typeof alignment !== 'string' || !ALIGNMENTS.includes(alignment)) throw new Error('invalid-formatting');
	const parsed = inspect(input), target = parsed.paragraphs[index];
	if (!target?.mark || target.prm !== 0 || /[\x00-\x1f]/u.test(parsed.text.slice(target.cpStart, target.cpEnd - 1))) throw new Error('unsupported-formatting');
	const row = target.row!;
	if (row.prls.some((p) => !SAFE.has(p.opcode))) throw new Error('opaque-formatting');
	const operands = alignmentOperands(row);
	if (!operands) throw new Error('inherited-formatting');
	if (operands.mirror && alignment !== 'center' && alignment !== 'justify') throw new Error('direction-dependent-alignment');
	const piece = parsed.pieces.find((p) => p.cpStart <= target.cpStart && p.cpEnd >= target.cpEnd);
	if (!piece || piece.prm !== 0) throw new Error('unsupported-formatting');
	const unit = piece.compressed ? 1 : 2, start = piece.fc + (target.cpStart - piece.cpStart) * unit;
	if (start !== row.start || start + (target.cpEnd - target.cpStart) * unit !== row.end ||
		parsed.paragraphs.some((p) => p !== target && p.row === row) ||
		parsed.pieces.some((p) => p !== piece && p.fc < row.end && row.start < p.fc + (p.cpEnd - p.cpStart) * (p.compressed ? 1 : 2))) throw new Error('shared-formatting');
	const extension = parsed.fib.fibRgFcLcbOffset + parsed.fib.fibRgFcLcbCount * 8;
	const fibEnd = extension + 2 + new DataView(parsed.word.buffer, parsed.word.byteOffset, parsed.word.byteLength).getUint16(extension, true) * 2;
	if (row.blob < fibEnd || parsed.rows.some((r) => r !== row && r.blob && r.blob < row.blobEnd && row.blob < r.blobEnd) ||
		parsed.chpx.pns.some((pn) => (pn & PAGE_NUMBER_MASK) === Math.floor(row.blob / 512)) ||
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
	const value = ALIGNMENTS.indexOf(alignment);
	if (operands.value.operand[0] === value) return input;
	const word = parsed.word.slice(); word[operands.value.at] = value;
	if (operands.mirror) word[operands.mirror.at] = value;
	return parsed.doc.rewrap(word, parsed.doc.tableBytes);
}
