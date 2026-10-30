/**
 * [MS-DOC] 2.9.? Prl / Sprm parsing shared by the character, paragraph, style
 * and list readers. A grpprl is a flat run of `Prl` records: a 16-bit sprm
 * opcode followed by an operand whose size is fixed by the opcode's `spra`
 * bits, or length-prefixed for the variable cases.
 *
 * @module doc-sprm
 */
import { DocResourceLimitError } from './ole-document-doc-pieces.js';

/** One decoded Prl: the opcode and a view of its operand bytes (without any length prefix). */
export interface DocRawSprm {
	readonly opcode: number;
	readonly operand: Uint8Array;
}

/** `sgc` (sprm group) of an opcode: 1 paragraph, 2 character, 3 picture, 4 section, 5 table. */
export function sprmGroup(opcode: number): number {
	return (opcode >>> 10) & 7;
}

/** Fixed operand sizes by `spra`; 6 is length-prefixed. */
const FIXED_SIZE: readonly number[] = [1, 1, 2, 4, 2, 2, -1, 3];

/** Shared processing budget for sprm parsing (decremented per Prl). */
export interface SprmBudget {
	remaining: number;
}

/** Thrown when the sprm budget is exhausted. */
export class DocSprmLimitError extends DocResourceLimitError {
	constructor() {
		super('DOC formatting SPRM count exceeds the processing limit');
		this.name = 'DocSprmLimitError';
	}
}

/**
 * Parse a grpprl. A truncated trailing record ends parsing (Word ignores it);
 * exceeding the budget throws {@link DocSprmLimitError}.
 */
export function parseGrpprl(bytes: Uint8Array, budget: SprmBudget): DocRawSprm[] {
	const out: DocRawSprm[] = [];
	let at = 0;
	while (at + 2 <= bytes.length) {
		if (budget.remaining-- <= 0) throw new DocSprmLimitError();
		const opcode = bytes[at]! | (bytes[at + 1]! << 8);
		const spra = opcode >>> 13;
		let start = at + 2;
		let size: number = FIXED_SIZE[spra]!;
		if (spra === 6) {
			if (opcode === 0xd608) {
				// sprmTDefTable: 16-bit cb that counts itself plus the data, minus one.
				if (start + 2 > bytes.length) break;
				size = (bytes[start]! | (bytes[start + 1]! << 8)) - 1;
				start += 2;
			} else {
				if (start + 1 > bytes.length) break;
				size = bytes[start]!;
				start += 1;
				if (opcode === 0xc615 && size === 255) size = chgTabsSize(bytes, start);
			}
		}
		if (size < 0 || start + size > bytes.length) break;
		out.push({ opcode, operand: bytes.subarray(start, start + size) });
		at = start + size;
	}
	return out;
}

/** Size of the oversized sprmPChgTabs operand (cb == 255): del/close arrays then add arrays. */
function chgTabsSize(bytes: Uint8Array, start: number): number {
	const del = bytes[start];
	if (del === undefined) return -1;
	const addAt = start + 1 + del * 4;
	const add = bytes[addAt];
	if (add === undefined) return -1;
	return 1 + del * 4 + 1 + add * 3;
}

/** Sprms superseded by a newer opcode for the same property; applied first so the modern one wins. */
const LEGACY = new Set<number>([
	0x2403, 0x840e, 0x840f, 0x8411, 0x6424, 0x6425, 0x6426, 0x6427, 0x6428, 0x6629, 0x442d, 0x2a42,
	0x486d, 0x486e, 0x4866, 0x6865,
]);

/** Order a grpprl so legacy duplicates of modern sprms are applied before them (stable otherwise). */
export function modernLast(sprms: readonly DocRawSprm[]): readonly DocRawSprm[] {
	if (!sprms.some((s) => LEGACY.has(s.opcode))) return sprms;
	return [...sprms.filter((s) => LEGACY.has(s.opcode)), ...sprms.filter((s) => !LEGACY.has(s.opcode))];
}

/** Little-endian readers over an operand. Return undefined when the operand is too short. */
export function u8(o: Uint8Array, at = 0): number | undefined {
	return at < o.length ? o[at] : undefined;
}
export function u16(o: Uint8Array, at = 0): number | undefined {
	return at + 2 <= o.length ? o[at]! | (o[at + 1]! << 8) : undefined;
}
export function i16(o: Uint8Array, at = 0): number | undefined {
	const v = u16(o, at);
	return v === undefined ? undefined : (v << 16) >> 16;
}
export function i32(o: Uint8Array, at = 0): number | undefined {
	return at + 4 <= o.length
		? (o[at]! | (o[at + 1]! << 8) | (o[at + 2]! << 16) | (o[at + 3]! << 24)) | 0
		: undefined;
}

const ICO = [
	undefined, '000000', '0000FF', '00FFFF', '00FF00', 'FF00FF', 'FF0000', 'FFFF00', 'FFFFFF',
	'000080', '008080', '008000', '800080', '800000', '808000', '808080', 'C0C0C0',
] as const;

/** Ico palette index (1..16) to `RRGGBB`; 0 (auto) and unknown values give undefined. */
export function docIcoToHex(ico: number): string | undefined {
	return ICO[ico];
}

/** COLORREF (r, g, b, fAuto) to `RRGGBB`; undefined for automatic colour. */
export function docColorRef(o: Uint8Array, at = 0): string | undefined {
	if (at + 4 > o.length || o[at + 3] === 0xff) return undefined;
	return [o[at]!, o[at + 1]!, o[at + 2]!].map((v) => v.toString(16).padStart(2, '0')).join('').toUpperCase();
}
