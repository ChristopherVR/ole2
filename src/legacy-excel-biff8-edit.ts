/** Fixed-length BIFF8 numeric edits. Every unrelated record and stream is preserved.
 * Formula tokens and cached results are preserved; this editor does not recalculate.
 * MS-XLS RkNumber, RK, MulRk and Number define the encodings used here.
 */
import { decodeRk, isSupportedBiff8Workbook, readRecords } from './legacy-excel-biff8.js';
import { unwrapXlsBytes } from './legacy-excel-cfb.js';

export type XlsNumericEditResult =
	| { status: 'edited'; bytes: Uint8Array; recalculationRequired: true }
	| { status: 'unchanged'; bytes: Uint8Array; reason: 'invalid-edit' | 'unsupported-workbook' | 'sheet-not-found' | 'cell-not-supported' | 'inexact-rk' | 'malformed-records' | 'container-not-writable' | 'ambiguous-cell' };

/** Encode without rounding. RK can store signed 30-bit integers, scaled integers,
 * or the top 30 bits of an IEEE double (optionally divided by 100).
 */
export function encodeRkExact(value: number): number | undefined {
	if (!Number.isFinite(value)) return undefined;
	for (const scale of [1, 100]) {
		const scaled = value * scale;
		if (Number.isInteger(scaled) && scaled >= -0x20000000 && scaled < 0x20000000) {
			const rk = ((scaled << 2) | 2 | (scale === 100 ? 1 : 0)) >>> 0;
			if (Object.is(decodeRk(rk), value)) return rk;
		}
		const view = new DataView(new ArrayBuffer(8));
		view.setFloat64(0, scaled, true);
		const rk = ((view.getUint32(4, true) & 0xfffffffc) | (scale === 100 ? 1 : 0)) >>> 0;
		if (Object.is(decodeRk(rk), value)) return rk;
	}
	return undefined;
}

/** Only top-level worksheet substreams: embedded chart BOF/EOF pairs nest. */
function worksheetRanges(bytes: Uint8Array): Array<{ start: number; end: number }> | undefined {
	const records = readRecords(bytes, 0, bytes.length);
	const last = records.at(-1);
	if (!last || bytes.subarray(last.dataOffset + last.length).some((b) => b !== 0)) return undefined;
	const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
	const ranges: Array<{ start: number; end: number }> = [];
	let depth = 0;
	let start: number | undefined;
	for (const record of records) {
		if (record.opcode === 0x0809) {
			if (record.length < 4 || view.getUint16(record.dataOffset, true) !== 0x0600 || depth >= 64) return undefined;
			if (depth === 0 && view.getUint16(record.dataOffset + 2, true) === 0x0010) start = record.headerOffset;
			depth++;
		} else if (record.opcode === 0x000a) {
			if (depth === 0 || record.length !== 0) return undefined;
			depth--;
			if (depth === 0 && start !== undefined) {
				ranges.push({ start, end: record.dataOffset });
				start = undefined;
			}
		}
	}
	if (depth !== 0) return undefined;
	// BoundSheet8 order is tab order; its lbPlyPos addresses the physical BOF.
	// MS-XLS 2.4.28: https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-xls/b9ec509a-235d-424e-871d-f8e721106501
	const globalsEnd = records.find((r) => r.opcode === 0x000a)?.headerOffset ?? 0;
	const bounds = records.filter((r) => r.opcode === 0x0085 && r.headerOffset < globalsEnd);
	// Retain compatibility with minimal bare BIFF streams having no tab directory.
	if (bounds.length === 0) return ranges;
	const ordered: Array<{ start: number; end: number }> = [];
	const used = new Set<number>();
	for (const bound of bounds) {
		if (bound.length < 8) return undefined;
		const p = bound.dataOffset;
		const nameLength = view.getUint8(p + 6);
		const width = (view.getUint8(p + 7) & 1) ? 2 : 1;
		if (nameLength < 1 || nameLength > 31 || bound.length !== 8 + nameLength * width) return undefined;
		const kind = view.getUint8(p + 5);
		if (![0, 1, 2, 6].includes(kind)) return undefined;
		if (kind !== 0) continue;
		const offset = view.getUint32(p, true);
		const range = ranges.find((r) => r.start === offset);
		if (!range || used.has(offset)) return undefined;
		used.add(offset);
		ordered.push(range);
	}
	return ordered.length === ranges.length ? ordered : undefined;
}

/** Edit an existing NUMBER/RK/MULRK cell in a selected worksheet (zero-based
 * worksheet index in BOUNDSHEET tab order, excluding macro/chart sheets). Unsupported edits return an explicit reason and original
 * bytes. No resize, formula execution, or cached formula recalculation occurs.
 */
export function editXlsNumericCell(
	input: Uint8Array,
	edit: { row: number; col: number; value: number; worksheetIndex?: number },
): XlsNumericEditResult {
	const unchanged = (reason: Extract<XlsNumericEditResult, { status: 'unchanged' }>['reason']): XlsNumericEditResult => ({ status: 'unchanged', bytes: input, reason });
	const sheet = edit.worksheetIndex ?? 0;
	if (!Number.isInteger(edit.row) || edit.row < 0 || edit.row > 65535 || !Number.isInteger(edit.col) || edit.col < 0 || edit.col > 255 || !Number.isInteger(sheet) || sheet < 0 || !Number.isFinite(edit.value)) return unchanged('invalid-edit');
	try {
		const { workbookBytes: bytes, rewrap } = unwrapXlsBytes(input);
		if (!isSupportedBiff8Workbook(bytes)) return unchanged('unsupported-workbook');
		const ranges = worksheetRanges(bytes);
		if (!ranges) return unchanged('malformed-records');
		const range = ranges[sheet];
		if (!range) return unchanged('sheet-not-found');
		const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
		let payload: number | undefined;
		let packed = false;
		let depth = 0;
		let matches = 0;
		for (const record of readRecords(bytes, range.start, range.end)) {
			if (record.opcode === 0x0809) { depth++; continue; }
			if (record.opcode === 0x000a) { depth--; continue; }
			if (depth !== 1) continue;
			const p = record.dataOffset;
			if (record.opcode === 0x00bd) {
				if (record.length < 12 || (record.length - 6) % 6 !== 0) return unchanged('malformed-records');
				const first = view.getUint16(p + 2, true);
				const lastCol = view.getUint16(p + record.length - 2, true);
				if (lastCol > 255 || lastCol < first || lastCol - first + 1 !== (record.length - 6) / 6) return unchanged('malformed-records');
				if (view.getUint16(p, true) === edit.row && edit.col >= first && edit.col <= lastCol) {
					if (++matches > 1) return unchanged('ambiguous-cell');
					payload = p + 6 + (edit.col - first) * 6;
					packed = true;
				}
			} else if (record.opcode === 0x0203 || record.opcode === 0x027e) {
				if (record.length !== (record.opcode === 0x0203 ? 14 : 10)) return unchanged('malformed-records');
				if (view.getUint16(p, true) === edit.row && view.getUint16(p + 2, true) === edit.col) {
					if (++matches > 1) return unchanged('ambiguous-cell');
					payload = p + 6;
					packed = record.opcode === 0x027e;
				}
			} else if ([0x0006, 0x0201, 0x0204, 0x0205, 0x00fd, 0x00d6].includes(record.opcode)) {
				if (record.length < 6) return unchanged('malformed-records');
				if (view.getUint16(p, true) === edit.row && view.getUint16(p + 2, true) === edit.col && ++matches > 1) return unchanged('ambiguous-cell');
			} else if (record.opcode === 0x00be) {
				if (record.length < 8 || (record.length - 6) % 2 !== 0) return unchanged('malformed-records');
				const first = view.getUint16(p + 2, true);
				const lastCol = view.getUint16(p + record.length - 2, true);
				if (lastCol > 255 || lastCol < first || lastCol - first + 1 !== (record.length - 6) / 2) return unchanged('malformed-records');
				if (view.getUint16(p, true) === edit.row && edit.col >= first && edit.col <= lastCol && ++matches > 1) return unchanged('ambiguous-cell');
			}
		}
		if (payload === undefined) return unchanged('cell-not-supported');
		const rk = packed ? encodeRkExact(edit.value) : undefined;
		if (packed && rk === undefined) return unchanged('inexact-rk');
		const out = bytes.slice();
		const outView = new DataView(out.buffer, out.byteOffset, out.byteLength);
		if (packed) outView.setUint32(payload, rk!, true);
		else outView.setFloat64(payload, edit.value, true);
		const result = rewrap ? rewrap(out) : out;
		if (result === input) return unchanged('container-not-writable');
		return { status: 'edited', bytes: result, recalculationRequired: true };
	} catch {
		return unchanged('malformed-records');
	}
}
