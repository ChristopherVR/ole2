/**
 * Table row properties (TAP) decoded from the table sprms of a row-end
 * paragraph ([MS-DOC] 2.6.3 table sprms, 2.9.321 TDefTableOperand,
 * 2.9.329 TC80, 2.9.25 Brc80). Opcodes not listed stay available as the
 * row's raw `sprms`.
 *
 * @module doc-table-props
 */
import { readBrc, readBrc80, readShd, readShd80 } from './doc-brc.js';
import type { DocBorder, DocShading } from './doc-brc.js';
import { findSprm, i16, u16, u8 } from './doc-sprm.js';
import type { DocRawSprm } from './doc-sprm.js';

const EMPTY = new Uint8Array(0);
const op = (s: DocRawSprm | undefined): Uint8Array => s?.operand ?? EMPTY;

export type DocCellVerticalAlign = 'top' | 'center' | 'bottom';

export interface DocCellProps {
	readonly widthTwips: number;
	/** `first` starts a horizontal merge, `continue` is swallowed by the cell on its left. */
	readonly horzMerge: 'first' | 'continue' | undefined;
	/** `restart` starts a vertical merge, `continue` is merged into the cell above. */
	readonly vertMerge: 'restart' | 'continue' | undefined;
	readonly verticalAlign: DocCellVerticalAlign;
	readonly noWrap: boolean;
	readonly borders: {
		readonly top?: DocBorder;
		readonly left?: DocBorder;
		readonly bottom?: DocBorder;
		readonly right?: DocBorder;
	};
	readonly shading: DocShading | undefined;
}

export interface DocRowProps {
	/** Cell boundary positions in twips; `cellCount + 1` entries. */
	readonly gridTwips: readonly number[];
	/** Table indent: left boundary plus the half cell gap, in twips. */
	readonly indentTwips: number;
	readonly gapHalfTwips: number;
	readonly alignment: 'left' | 'center' | 'right';
	readonly isHeader: boolean;
	readonly cantSplit: boolean;
	/** Row height in twips; 0 means automatic. */
	readonly heightTwips: number;
	readonly heightRule: 'auto' | 'atLeast' | 'exact';
	readonly borders: {
		readonly top?: DocBorder;
		readonly left?: DocBorder;
		readonly bottom?: DocBorder;
		readonly right?: DocBorder;
		readonly insideH?: DocBorder;
		readonly insideV?: DocBorder;
	};
	readonly cells: readonly DocCellProps[];
	readonly sprms: readonly DocRawSprm[];
}

interface MutableCell {
	widthTwips: number;
	horzMerge: 'first' | 'continue' | undefined;
	vertMerge: 'restart' | 'continue' | undefined;
	verticalAlign: DocCellVerticalAlign;
	noWrap: boolean;
	borders: { top?: DocBorder; left?: DocBorder; bottom?: DocBorder; right?: DocBorder };
	shading: DocShading | undefined;
}

/** Decode the TAP of a row-end paragraph from its table sprms. Returns undefined without a sprmTDefTable. */
export function readRowProps(sprms: readonly DocRawSprm[]): DocRowProps | undefined {
	const def = findSprm(sprms, 0xd608);
	if (!def || def.operand.length < 1) return undefined;
	const o = def.operand;
	const itcMac = o[0]!;
	const dv = new DataView(o.buffer, o.byteOffset, o.byteLength);
	const gridTwips: number[] = [];
	for (let i = 0; i <= itcMac && 1 + i * 2 + 2 <= o.length; i++) gridTwips.push(dv.getInt16(1 + i * 2, true));
	const tcStart = 1 + (itcMac + 1) * 2;
	const cells: MutableCell[] = [];
	for (let i = 0; i < itcMac; i++) {
		const at = tcStart + i * 20;
		const g1 = gridTwips[i];
		const g2 = gridTwips[i + 1];
		const cell: MutableCell = {
			widthTwips: g1 !== undefined && g2 !== undefined ? g2 - g1 : 0,
			horzMerge: undefined,
			vertMerge: undefined,
			verticalAlign: 'top',
			noWrap: false,
			borders: {},
			shading: undefined,
		};
		if (at + 20 <= o.length) {
			const tcgrf = o[at]! | (o[at + 1]! << 8);
			cell.horzMerge = (tcgrf & 3) === 1 ? 'first' : (tcgrf & 3) === 2 ? 'continue' : undefined;
			const v = (tcgrf >> 5) & 3;
			cell.vertMerge = v === 1 ? 'continue' : v === 3 ? 'restart' : undefined;
			cell.verticalAlign = (['top', 'center', 'bottom'] as const)[(tcgrf >> 7) & 3] ?? 'top';
			cell.noWrap = (tcgrf & 0x2000) !== 0;
			const top = readBrc80(o, at + 4);
			const left = readBrc80(o, at + 8);
			const bottom = readBrc80(o, at + 12);
			const right = readBrc80(o, at + 16);
			cell.borders = { ...(top && { top }), ...(left && { left }), ...(bottom && { bottom }), ...(right && { right }) };
		}
		cells.push(cell);
	}
	applyCellSprms(sprms, cells);
	const height = i16(op(findSprm(sprms, 0x9407))) ?? 0;
	const gapHalf = i16(op(findSprm(sprms, 0x9602))) ?? 0;
	const jc = u16(op(findSprm(sprms, 0x5400))) ?? 0;
	return {
		gridTwips,
		indentTwips: (gridTwips[0] ?? 0) + gapHalf,
		gapHalfTwips: gapHalf,
		alignment: jc === 1 ? 'center' : jc === 2 ? 'right' : 'left',
		isHeader: (u8(op(findSprm(sprms, 0x3404))) ?? 0) !== 0,
		cantSplit: (u8(op(findSprm(sprms, 0x3403))) ?? 0) !== 0,
		heightTwips: Math.abs(height),
		heightRule: height === 0 ? 'auto' : height > 0 ? 'atLeast' : 'exact',
		borders: readTableBorders(sprms),
		cells,
		sprms,
	};
}

function readTableBorders(sprms: readonly DocRawSprm[]): DocRowProps['borders'] {
	const out: { -readonly [K in keyof DocRowProps['borders']]: DocRowProps['borders'][K] } = {};
	const names = ['top', 'left', 'bottom', 'right', 'insideH', 'insideV'] as const;
	const modern = findSprm(sprms, 0xd613);
	const legacy = findSprm(sprms, 0xd605);
	for (let i = 0; i < 6; i++) {
		let b: DocBorder | undefined;
		if (modern && modern.operand.length >= 6 * 8) b = readBrc(modern.operand, i * 8);
		else if (legacy && legacy.operand.length >= 24) b = readBrc80(legacy.operand, i * 4);
		if (b) out[names[i]!] = b;
	}
	return out;
}

const SIDES = ['top', 'left', 'bottom', 'right'] as const;

function applyCellSprms(sprms: readonly DocRawSprm[], cells: MutableCell[]): void {
	const range = (first: number, lim: number) => cells.slice(Math.max(0, first), Math.min(cells.length, lim));
	for (const s of sprms) {
		const o = s.operand;
		switch (s.opcode) {
			case 0xd609: // sprmTDefTableShd: one Shd80 per cell
				for (let i = 0; i < cells.length && i * 2 + 2 <= o.length; i++) cells[i]!.shading = readShd80(o, i * 2);
				break;
			case 0x7627: // sprmTSetShd: itcFirst, itcLim, Shd80
				for (const c of range(o[0]!, o[1]!)) c.shading = readShd80(o, 2);
				break;
			case 0xd620: // sprmTSetBrc: itcFirst, itcLim, grfbrc, Brc80
				for (const c of range(o[0]!, o[1]!)) {
					const b = readBrc80(o, 3);
					const g = o[2]!;
					SIDES.forEach((side, k) => {
						if (g & (1 << k)) setBorder(c, side, b);
					});
				}
				break;
			case 0xd62b: // sprmTVertMerge: itc, flags
				if (cells[o[0]!]) cells[o[0]!]!.vertMerge = o[1] === 1 ? 'continue' : o[1] === 3 ? 'restart' : undefined;
				break;
			case 0xd62c: // sprmTVertAlign: itcFirst, itcLim, value
				for (const c of range(o[0]!, o[1]!)) c.verticalAlign = (['top', 'center', 'bottom'] as const)[o[2]!] ?? 'top';
				break;
			case 0xd635: // sprmTCellWidth: itcFirst, itcLim, width
				for (const c of range(o[0]!, o[1]!)) c.widthTwips = o[2]! | (o[3]! << 8);
				break;
			default:
		}
	}
	for (const s of sprms) {
		// Word 2002+ writes 10-byte SHD records for exact colours; they override the 16-colour Shd80.
		if (s.opcode === 0xd612) {
			for (let i = 0; i < cells.length && i * 10 + 10 <= s.operand.length; i++) {
				const shd = readShd(s.operand, i * 10);
				if (shd) cells[i]!.shading = shd;
			}
		}
	}
}

function setBorder(c: MutableCell, side: (typeof SIDES)[number], b: DocBorder | undefined): void {
	if (b) c.borders[side] = b;
	else delete c.borders[side];
}
