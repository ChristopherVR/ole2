/** Preserve main-story field records while shifting positions for a plain-text edit.
 * [MS-DOC] 2.8.25 Plcfld and 2.9.88/89 Fld/fldch. */
import type { FcLcb } from './ole-document-doc-fib.js';
import { DocResourceLimitError } from './ole-document-doc-pieces.js';

export class UnsupportedDocFieldEditError extends Error {
	constructor(message: string) {
		super(message);
		this.name = 'UnsupportedDocFieldEditError';
	}
}

/** Verify nesting and refuse editing any field code/result range, including
 * XE/TC/RD/TA/PRIVATE fields deliberately omitted from the field PLC. */
function checkFieldRanges(markers: Iterable<{ cp: number; ch: number }>, editStart: number, editEnd: number, maxDepth: number): void {
	const stack: Array<{ start: number; separated: boolean }> = [];
	for (const { cp, ch } of markers) {
		if (ch === 0x13) {
			if (stack.length >= maxDepth) throw new DocResourceLimitError('DOC field nesting exceeds the processing limit');
			stack.push({ start: cp, separated: false });
		} else if (ch === 0x14) {
			const field = stack.at(-1);
			if (!field || field.separated) throw new Error('Invalid DOC field separator nesting');
			field.separated = true;
		} else if (ch === 0x15) {
			const field = stack.pop();
			if (!field) throw new Error('Unmatched DOC field end');
			if (field.start < editEnd && cp + 1 > editStart)
				throw new UnsupportedDocFieldEditError('Paragraph overlaps a field code/result range');
		}
	}
	if (stack.length) throw new Error('Unterminated DOC field');
}

function* bodyMarkers(body: string): Generator<{ cp: number; ch: number }> {
	for (let cp = 0; cp < body.length; cp++) {
		const ch = body.charCodeAt(cp);
		if (ch >= 0x13 && ch <= 0x15) yield { cp, ch };
	}
}

/** Return a same-size field PLC with shifted actual CPs and opaque Fld records.
 * The undefined sentinel remains sorted/largest; it need not equal the story end.
 * The source table and every field code/result character are left unchanged. */
export function shiftDocMainFieldCps(
	table: Uint8Array,
	at: FcLcb,
	body: string,
	editStart: number,
	editEnd: number,
	delta: number,
	maxFields: number,
): Uint8Array | undefined {
	if (![editStart, editEnd, delta, maxFields].every(Number.isSafeInteger) ||
		editStart < 0 || editEnd <= editStart || editEnd > body.length || maxFields < 0 || body.length + delta < 0)
		throw new Error('Invalid field edit bounds or budget');
	checkFieldRanges(bodyMarkers(body), editStart, editEnd, maxFields);
	if (at.lcb === 0) return undefined;
	if (!Number.isSafeInteger(at.fc) || at.fc < 0 || at.lcb < 4 || (at.lcb - 4) % 6 !== 0 || at.fc > table.length - at.lcb)
		throw new Error('Invalid main-story field PLC bounds');
	const n = (at.lcb - 4) / 6;
	if (n > maxFields) throw new DocResourceLimitError('DOC field count exceeds the processing limit');
	const source = table.subarray(at.fc, at.fc + at.lcb);
	const view = new DataView(source.buffer, source.byteOffset, source.byteLength);
	const markers: Array<{ cp: number; ch: number }> = [];
	let previous = -1;
	for (let i = 0; i < n; i++) {
		const cp = view.getInt32(i * 4, true);
		const ch = source[(n + 1) * 4 + i * 2]! & 0x1f;
		if (cp <= previous || cp >= body.length || ![0x13, 0x14, 0x15].includes(ch) || body.charCodeAt(cp) !== ch)
			throw new Error('Invalid main-story field marker');
		markers.push({ cp, ch });
		previous = cp;
	}
	checkFieldRanges(markers, editStart, editEnd, maxFields);
	const sentinel = view.getInt32(n * 4, true);
	if (sentinel < 0 || sentinel <= previous) throw new Error('Invalid field PLC sentinel');
	const result = source.slice();
	const out = new DataView(result.buffer);
	let shiftedLast = -1;
	for (let i = 0; i < n; i++) {
		const cp = markers[i]!.cp;
		const shifted = cp >= editEnd ? cp + delta : cp;
		if (shifted < 0 || shifted >= body.length + delta || shifted <= shiftedLast)
			throw new Error('Field position shifts outside the main story');
		out.setInt32(i * 4, shifted, true);
		shiftedLast = shifted;
	}
	let shiftedSentinel = sentinel >= editEnd ? sentinel + delta : sentinel;
	// Undefined sentinels near INT32_MAX must not overflow. Retain the original
	// sentinel when it still sorts after every actual field character.
	if (shiftedSentinel > 0x7fffffff || shiftedSentinel <= shiftedLast) shiftedSentinel = sentinel;
	if (shiftedSentinel <= shiftedLast) throw new Error('No representable field PLC sentinel');
	out.setInt32(n * 4, shiftedSentinel, true);
	return result;
}
