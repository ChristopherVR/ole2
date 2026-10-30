/**
 * Table row properties (TAP) decoded from the table sprms of a row-end
 * paragraph ([MS-DOC] 2.6.3 table sprms, 2.9.321 TDefTableOperand,
 * 2.9.329 TC80, 2.9.25 Brc80). Opcodes not listed stay available as the
 * row's raw `sprms`.
 *
 * @module doc-table-props
 */
import { lastSprm, operandI16, operandU16, operandU8, type DocStructureSprm } from './doc-sprm-reader.js';

export interface DocBorder {
	/** Readable border style; `other` when the numeric `type` is not mapped. */
	readonly style: string;
	/** Raw brcType. */
	readonly type: number;
	/** Line width in 1/8 pt. */
	readonly widthEighthPt: number;
	/** `RRGGBB` hex or `auto`. */
	readonly color: string;
	readonly spacePt: number;
	readonly shadow: boolean;
	readonly frame: boolean;
}

export type DocCellVerticalAlign = 'top' | 'center' | 'bottom';

export interface DocCellShading {
	/** Shading pattern (ipat): 0 is clear (the background is the fill), 1 is solid (the foreground is the fill). */
	readonly pattern: number;
	/** Foreground `RRGGBB` or `auto`. */
	readonly foreground: string;
	/** Background `RRGGBB` or `auto`; the fill colour for solid patterns. */
	readonly background: string;
}

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
	readonly shading: DocCellShading | undefined;
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
	readonly sprms: readonly DocStructureSprm[];
}

const ICO = ['auto', '000000', '0000FF', '00FFFF', '00FF00', 'FF00FF', 'FF0000', 'FFFF00', 'FFFFFF', '000080', '008080', '008000', '800080', '800000', '808000', '808080', 'C0C0C0'];
const BORDER_STYLES: Record<number, string> = {
	0: 'none', 1: 'single', 2: 'thick', 3: 'double', 5: 'hairline', 6: 'dotted', 7: 'dashed', 8: 'dotDash', 9: 'dotDotDash', 10: 'triple',
	11: 'thinThickSmallGap', 12: 'thickThinSmallGap', 13: 'thinThickThinSmallGap', 14: 'thinThickMediumGap', 15: 'thickThinMediumGap',
	16: 'thinThickThinMediumGap', 17: 'thinThickLargeGap', 18: 'thickThinLargeGap', 19: 'thinThickThinLargeGap', 20: 'wave', 21: 'doubleWave',
	22: 'dashSmallGap', 23: 'dashDotStroked', 24: 'threeDEmboss', 25: 'threeDEngrave', 26: 'outset', 27: 'inset', 255: 'none',
};

const hex = (n: number) => n.toString(16).toUpperCase().padStart(2, '0');
function colorRef(b: Uint8Array, at: number): string {
	// COLORREF: red, green, blue, then an auto marker (0xFF) in the high byte.
	return b[at + 3] === 0xff ? 'auto' : hex(b[at]!) + hex(b[at + 1]!) + hex(b[at + 2]!);
}

/** Decode a 4-byte Brc80 at `at`. */
export function readBrc80(b: Uint8Array, at: number): DocBorder | undefined {
	if (at + 4 > b.length) return undefined;
	const type = b[at + 1]!;
	const width = b[at]!;
	if (type === 0xff || (type === 0 && width === 0)) return undefined;
	return {
		style: BORDER_STYLES[type] ?? 'other',
		type,
		widthEighthPt: width,
		color: ICO[b[at + 2]!] ?? 'auto',
		spacePt: b[at + 3]! & 0x1f,
		shadow: (b[at + 3]! & 0x20) !== 0,
		frame: (b[at + 3]! & 0x40) !== 0,
	};
}

function readShd80(v: number): DocCellShading | undefined {
	const fore = v & 0x1f;
	const back = (v >> 5) & 0x1f;
	const pattern = (v >> 10) & 0x3f;
	if (v === 0) return undefined;
	return { pattern, foreground: ICO[fore] ?? 'auto', background: ICO[back] ?? 'auto' };
}

function readShd(b: Uint8Array, at: number): DocCellShading | undefined {
	// SHD: cvFore (4), cvBack (4), ipat (2).
	if (at + 10 > b.length) return undefined;
	const pattern = b[at + 8]! | (b[at + 9]! << 8);
	const foreground = colorRef(b, at);
	const background = colorRef(b, at + 4);
	if (pattern === 0xffff) return undefined;
	return { pattern, foreground, background };
}

interface MutableCell {
	widthTwips: number;
	horzMerge: 'first' | 'continue' | undefined;
	vertMerge: 'restart' | 'continue' | undefined;
	verticalAlign: DocCellVerticalAlign;
	noWrap: boolean;
	borders: { top?: DocBorder; left?: DocBorder; bottom?: DocBorder; right?: DocBorder };
	shading: DocCellShading | undefined;
}

/** Decode the TAP of a row-end paragraph from its table sprms. Returns undefined without a sprmTDefTable. */
export function readRowProps(sprms: readonly DocStructureSprm[]): DocRowProps | undefined {
	const def = lastSprm(sprms, 0xd608);
	if (!def || def.operand.length < 3) return undefined;
	const o = def.operand;
	const itcMac = o[2]!;
	const dv = new DataView(o.buffer, o.byteOffset, o.byteLength);
	const gridTwips: number[] = [];
	for (let i = 0; i <= itcMac && 3 + i * 2 + 2 <= o.length; i++) gridTwips.push(dv.getInt16(3 + i * 2, true));
	const tcStart = 3 + (itcMac + 1) * 2;
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
	const height = operandI16(lastSprm(sprms, 0x9407)) ?? 0;
	const gapHalf = operandI16(lastSprm(sprms, 0x9602)) ?? 0;
	const jc = operandU16(lastSprm(sprms, 0x5400)) ?? 0;
	const borders = readTableBorders(sprms);
	return {
		gridTwips,
		indentTwips: (gridTwips[0] ?? 0) + gapHalf,
		gapHalfTwips: gapHalf,
		alignment: jc === 1 ? 'center' : jc === 2 ? 'right' : 'left',
		isHeader: (operandU8(lastSprm(sprms, 0x3404)) ?? 0) !== 0,
		cantSplit: (operandU8(lastSprm(sprms, 0x3403)) ?? 0) !== 0,
		heightTwips: Math.abs(height),
		heightRule: height === 0 ? 'auto' : height > 0 ? 'atLeast' : 'exact',
		borders,
		cells,
		sprms,
	};
}

function readTableBorders(sprms: readonly DocStructureSprm[]): DocRowProps['borders'] {
	const out: { -readonly [K in keyof DocRowProps['borders']]: DocRowProps['borders'][K] } = {};
	const names = ['top', 'left', 'bottom', 'right', 'insideH', 'insideV'] as const;
	const modern = lastSprm(sprms, 0xd613);
	const legacy = lastSprm(sprms, 0xd605);
	for (let i = 0; i < 6; i++) {
		let b: DocBorder | undefined;
		if (modern && modern.operand.length >= 1 + 6 * 8) b = readBrc(modern.operand, 1 + i * 8);
		else if (legacy && legacy.operand.length >= 1 + 24) b = readBrc80(legacy.operand, 1 + i * 4);
		if (b) out[names[i]!] = b;
	}
	return out;
}

/** Decode an 8-byte Brc (COLORREF, width, type, space/flags). */
function readBrc(o: Uint8Array, at: number): DocBorder | undefined {
	if (at + 8 > o.length) return undefined;
	const width = o[at + 4]!;
	const type = o[at + 5]!;
	if (type === 0xff || (type === 0 && width === 0)) return undefined;
	return {
		style: BORDER_STYLES[type] ?? 'other',
		type,
		widthEighthPt: width,
		color: colorRef(o, at),
		spacePt: o[at + 6]! & 0x1f,
		shadow: (o[at + 6]! & 0x20) !== 0,
		frame: (o[at + 6]! & 0x40) !== 0,
	};
}

function applyCellSprms(sprms: readonly DocStructureSprm[], cells: MutableCell[]): void {
	const range = (first: number, lim: number) => cells.slice(Math.max(0, first), Math.min(cells.length, lim));
	for (const s of sprms) {
		const o = s.operand;
		switch (s.opcode) {
			case 0xd609: // sprmTDefTableShd: one Shd80 per cell
				for (let i = 0; i < cells.length && 1 + i * 2 + 2 <= o.length; i++) cells[i]!.shading = readShd80(o[1 + i * 2]! | (o[2 + i * 2]! << 8));
				break;
			case 0x7627: // sprmTSetShd: itcFirst, itcLim, Shd80
				for (const c of range(o[0]!, o[1]!)) c.shading = readShd80(o[2]! | (o[3]! << 8));
				break;
			case 0xd620: // sprmTSetBrc: itcFirst, itcLim, grfbrc, Brc80
				for (const c of range(o[1]!, o[2]!)) {
					const b = readBrc80(o, 4);
					const g = o[3]!;
					if (g & 1) setBorder(c, 'top', b);
					if (g & 2) setBorder(c, 'left', b);
					if (g & 4) setBorder(c, 'bottom', b);
					if (g & 8) setBorder(c, 'right', b);
				}
				break;
			case 0xd62b: // sprmTVertMerge: itc, flags
				if (cells[o[1]!]) cells[o[1]!]!.vertMerge = o[2] === 1 ? 'continue' : o[2] === 3 ? 'restart' : undefined;
				break;
			case 0xd62c: // sprmTVertAlign: itcFirst, itcLim, value
				for (const c of range(o[1]!, o[2]!)) c.verticalAlign = (['top', 'center', 'bottom'] as const)[o[3]!] ?? 'top';
				break;
			case 0xd635: // sprmTCellWidth: itcFirst, itcLim, width
				for (const c of range(o[1]!, o[2]!)) c.widthTwips = o[3]! | (o[4]! << 8);
				break;
			default:
		}
	}
	for (const s of sprms) {
		// Word 2002+ writes 10-byte SHD records for exact colours; they override the 16-colour Shd80.
		if (s.opcode === 0xd612) {
			const o = s.operand;
			for (let i = 0; i < cells.length && 1 + i * 10 + 10 <= o.length; i++) {
				const shd = readShd(o, 1 + i * 10);
				if (shd) cells[i]!.shading = shd;
			}
		}
	}
}

function setBorder(c: MutableCell, side: 'top' | 'left' | 'bottom' | 'right', b: DocBorder | undefined): void {
	if (b) c.borders[side] = b;
	else delete c.borders[side];
}
