/**
 * Read-only [MS-DOC] 2.6 property modifier (`Prl`/sprm) group parsing shared by
 * the structure readers (tables, sections). Operands stay opaque bytes; the
 * callers decode only the opcodes they own.
 *
 * @module doc-sprm-reader
 */
import { DocResourceLimitError } from './ole-document-doc-pieces.js';

/** One sprm: opcode and its operand bytes. */
export interface DocStructureSprm {
	readonly opcode: number;
	/** Operand bytes. For variable-size sprms this includes the size prefix. */
	readonly operand: Uint8Array;
}

/** Shared sprm count budget across one structure read. */
export interface SprmBudget {
	remaining: number;
}

const FIXED_OPERAND = [1, 1, 2, 4, 2, 2, 0, 3] as const;

/** Sprm group of an opcode: 1 paragraph, 2 character, 3 picture, 4 section, 5 table. */
export function sprmGroup(opcode: number): number {
	return (opcode >>> 10) & 7;
}

/** Parse the complete `GrpPrl` in `[start, end)`; stops (without throwing) at a truncated tail. */
export function readGrpprl(data: Uint8Array, start: number, end: number, budget: SprmBudget): DocStructureSprm[] {
	const out: DocStructureSprm[] = [];
	for (let at = start; at + 2 <= end; ) {
		if (budget.remaining <= 0) throw new DocResourceLimitError('DOC sprm count exceeds the processing limit');
		budget.remaining--;
		const opcode = data[at]! | (data[at + 1]! << 8);
		const spra = opcode >>> 13;
		let size = FIXED_OPERAND[spra]!;
		if (spra === 6) {
			if (opcode === 0xd608) {
				// sprmTDefTable: a 16-bit size that counts itself and the bytes after it, minus one.
				if (at + 4 > end) break;
				size = (data[at + 2]! | (data[at + 3]! << 8)) + 1;
			} else if (opcode === 0xc615 && data[at + 2] === 255) {
				// sprmPChgTabs with cb == 255: derive the length from its two counts.
				const del = data[at + 3] ?? 0;
				const add = data[at + 4 + del * 4] ?? 0;
				size = 1 + 1 + del * 4 + 1 + add * 3;
			} else {
				if (at + 3 > end) break;
				size = 1 + data[at + 2]!;
			}
		}
		if (at + 2 + size > end) break;
		out.push({ opcode, operand: data.subarray(at + 2, at + 2 + size) });
		at += 2 + size;
	}
	return out;
}

/** Last sprm with `opcode`, as later sprms override earlier ones. */
export function lastSprm(sprms: readonly DocStructureSprm[], opcode: number): DocStructureSprm | undefined {
	for (let i = sprms.length - 1; i >= 0; i--) if (sprms[i]!.opcode === opcode) return sprms[i];
	return undefined;
}

export function operandU8(sprm: DocStructureSprm | undefined): number | undefined {
	return sprm && sprm.operand.length >= 1 ? sprm.operand[0] : undefined;
}

export function operandU16(sprm: DocStructureSprm | undefined): number | undefined {
	return sprm && sprm.operand.length >= 2 ? sprm.operand[0]! | (sprm.operand[1]! << 8) : undefined;
}

export function operandI16(sprm: DocStructureSprm | undefined): number | undefined {
	const v = operandU16(sprm);
	return v === undefined ? undefined : (v << 16) >> 16;
}

export function operandI32(sprm: DocStructureSprm | undefined): number | undefined {
	if (!sprm || sprm.operand.length < 4) return undefined;
	const o = sprm.operand;
	return o[0]! | (o[1]! << 8) | (o[2]! << 16) | (o[3]! << 24);
}
