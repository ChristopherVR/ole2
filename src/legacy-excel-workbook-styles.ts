/**
 * BIFF8 formatting records: `FONT`, `FORMAT`, `XF`, `XFEXT` and `PALETTE`,
 * plus the default 64-colour palette and Excel's built-in (en-US) number
 * format codes for ids that a workbook does not redefine with `FORMAT`.
 *
 * @module legacy-excel-workbook-styles
 */
import { ByteReader } from './legacy-excel-workbook-records.js';
import type {
	XlsBorderEdge,
	XlsBorderStyle,
	XlsColor,
	XlsFillPattern,
	XlsFont,
	XlsHorizontalAlign,
	XlsUnderline,
	XlsVerticalAlign,
	XlsXf,
} from './legacy-excel-workbook-types.js';

/** The BIFF8 default palette: indices 0-7 are fixed, 8-63 can be redefined by `PALETTE`. */
export const DEFAULT_XLS_PALETTE: readonly string[] = (
	'000000 FFFFFF FF0000 00FF00 0000FF FFFF00 FF00FF 00FFFF ' +
	'000000 FFFFFF FF0000 00FF00 0000FF FFFF00 FF00FF 00FFFF 800000 008000 000080 808000 800080 ' +
	'008080 C0C0C0 808080 9999FF 993366 FFFFCC CCFFFF 660066 FF8080 0066CC CCCCFF 000080 FF00FF ' +
	'FFFF00 00FFFF 800080 800000 008080 0000FF 00CCFF CCFFFF CCFFCC FFFF99 99CCFF FF99CC CC99FF ' +
	'FFCC99 3366FF 33CCCC 99CC00 FFCC00 FF9900 FF6600 666699 969696 003366 339966 003300 333300 ' +
	'993300 993366 333399 333333'
).split(' ');

/** Built-in number formats by id (the en-US codes Excel uses when no `FORMAT` record overrides them). */
export const XLS_BUILTIN_FORMATS: Readonly<Record<number, string>> = {
	0: 'General',
	1: '0',
	2: '0.00',
	3: '#,##0',
	4: '#,##0.00',
	5: '"$"#,##0_);\\("$"#,##0\\)',
	6: '"$"#,##0_);[Red]\\("$"#,##0\\)',
	7: '"$"#,##0.00_);\\("$"#,##0.00\\)',
	8: '"$"#,##0.00_);[Red]\\("$"#,##0.00\\)',
	9: '0%',
	10: '0.00%',
	11: '0.00E+00',
	12: '# ?/?',
	13: '# ??/??',
	14: 'm/d/yyyy',
	15: 'd-mmm-yy',
	16: 'd-mmm',
	17: 'mmm-yy',
	18: 'h:mm AM/PM',
	19: 'h:mm:ss AM/PM',
	20: 'h:mm',
	21: 'h:mm:ss',
	22: 'm/d/yyyy h:mm',
	37: '#,##0 ;(#,##0)',
	38: '#,##0 ;[Red](#,##0)',
	39: '#,##0.00;(#,##0.00)',
	40: '#,##0.00;[Red](#,##0.00)',
	41: '_(* #,##0_);_(* \\(#,##0\\);_(* "-"_);_(@_)',
	42: '_("$"* #,##0_);_("$"* \\(#,##0\\);_("$"* "-"_);_(@_)',
	43: '_(* #,##0.00_);_(* \\(#,##0.00\\);_(* "-"??_);_(@_)',
	44: '_("$"* #,##0.00_);_("$"* \\(#,##0.00\\);_("$"* "-"??_);_(@_)',
	45: 'mm:ss',
	46: '[h]:mm:ss',
	47: 'mm:ss.0',
	48: '##0.0E+0',
	49: '@',
};

/** Resolve a BIFF colour index against the palette (system and automatic colours stay `auto`). */
export function paletteColor(index: number, palette: readonly string[]): XlsColor {
	const rgb = index < 64 ? palette[index] : undefined;
	return rgb ? { index, rgb } : { index, auto: true };
}

const UNDERLINE: Readonly<Record<number, XlsUnderline>> = {
	0x01: 'single',
	0x02: 'double',
	0x21: 'singleAccounting',
	0x22: 'doubleAccounting',
};

/** `FONT`: height in twips, attributes, colour index, weight, script, underline, family, name. */
export function parseFont(data: Uint8Array, palette: readonly string[]): XlsFont {
	const r = new ByteReader(data);
	const height = r.u16();
	const flags = r.u16();
	const color = r.u16();
	const weight = r.u16();
	const script = r.u16();
	const underline = r.u8();
	const family = r.u8();
	const charset = r.u8();
	r.skip(1);
	return {
		name: r.shortXlString(),
		size: height / 20,
		bold: weight >= 600,
		weight,
		italic: (flags & 0x02) !== 0,
		underline: UNDERLINE[underline] ?? 'none',
		strike: (flags & 0x08) !== 0,
		outline: (flags & 0x10) !== 0,
		shadow: (flags & 0x20) !== 0,
		script: script === 1 ? 'superscript' : script === 2 ? 'subscript' : 'none',
		color: paletteColor(color, palette),
		family,
		charset,
	};
}

/** `FORMAT`: number format id and its code. */
export function parseFormat(data: Uint8Array): { id: number; code: string } {
	const r = new ByteReader(data);
	const id = r.u16();
	return { id, code: r.xlString() };
}

/** `PALETTE`: colours for indices 8 onwards. */
export function parsePalette(data: Uint8Array): string[] {
	const r = new ByteReader(data);
	const count = r.u16();
	const palette = [...DEFAULT_XLS_PALETTE];
	for (let i = 0; i < count && i + 8 < 64; i++) {
		const rgb = [r.u8(), r.u8(), r.u8()];
		r.skip(1);
		palette[i + 8] = rgb.map((c) => c.toString(16).padStart(2, '0').toUpperCase()).join('');
	}
	return palette;
}

const H_ALIGN: readonly XlsHorizontalAlign[] = [
	'general',
	'left',
	'center',
	'right',
	'fill',
	'justify',
	'centerContinuous',
	'distributed',
];
const V_ALIGN: readonly XlsVerticalAlign[] = ['top', 'center', 'bottom', 'justify', 'distributed'];
const BORDER: readonly (XlsBorderStyle | undefined)[] = [
	undefined,
	'thin',
	'medium',
	'dashed',
	'dotted',
	'thick',
	'double',
	'hair',
	'mediumDashed',
	'dashDot',
	'mediumDashDot',
	'dashDotDot',
	'mediumDashDotDot',
	'slantDashDot',
];
const PATTERN: readonly XlsFillPattern[] = [
	'none',
	'solid',
	'mediumGray',
	'darkGray',
	'lightGray',
	'darkHorizontal',
	'darkVertical',
	'darkDown',
	'darkUp',
	'darkGrid',
	'darkTrellis',
	'lightHorizontal',
	'lightVertical',
	'lightDown',
	'lightUp',
	'lightGrid',
	'lightTrellis',
	'gray125',
	'gray0625',
];

function edge(style: number, color: number, palette: readonly string[]): XlsBorderEdge | undefined {
	const name = BORDER[style];
	return name ? { style: name, color: paletteColor(color, palette) } : undefined;
}

/** `XF` (20 bytes): font, number format, protection, alignment, borders and fill. */
export function parseXf(
	data: Uint8Array,
	palette: readonly string[],
	formatCode: (id: number) => string,
): XlsXf {
	const r = new ByteReader(data);
	const fontIndex = r.u16();
	const numFmtId = r.u16();
	const type = r.u16();
	const align = r.u8();
	const rotation = r.u8();
	const indentFlags = r.u8();
	r.skip(1);
	const border1 = r.u32();
	const border2 = r.u32();
	const fill = r.u16();
	const diag = border1 >>> 30;
	const xf: XlsXf = {
		font: fontIndex >= 4 ? fontIndex - 1 : fontIndex,
		numFmtId,
		numFmt: formatCode(numFmtId),
		isStyle: (type & 0x04) !== 0,
		parent: type >>> 4,
		locked: (type & 0x01) !== 0,
		hidden: (type & 0x02) !== 0,
		alignment: {
			horizontal: H_ALIGN[align & 0x07] ?? 'general',
			vertical: V_ALIGN[(align >> 4) & 0x07] ?? 'bottom',
			wrap: (align & 0x08) !== 0,
			shrinkToFit: (indentFlags & 0x10) !== 0,
			indent: indentFlags & 0x0f,
			rotation,
			readingOrder: (indentFlags >> 6) & 0x03,
		},
		border: { diagonalDown: (diag & 1) !== 0, diagonalUp: (diag & 2) !== 0 },
		fill: {
			pattern: PATTERN[(border2 >>> 26) & 0x3f] ?? 'none',
			fg: paletteColor(fill & 0x7f, palette),
			bg: paletteColor((fill >> 7) & 0x7f, palette),
		},
	};
	const edges: [keyof XlsXf['border'], XlsBorderEdge | undefined][] = [
		['left', edge(border1 & 0x0f, (border1 >>> 16) & 0x7f, palette)],
		['right', edge((border1 >>> 4) & 0x0f, (border1 >>> 23) & 0x7f, palette)],
		['top', edge((border1 >>> 8) & 0x0f, border2 & 0x7f, palette)],
		['bottom', edge((border1 >>> 12) & 0x0f, (border2 >>> 7) & 0x7f, palette)],
		[
			'diagonal',
			diag ? edge((border2 >>> 21) & 0x0f, (border2 >>> 14) & 0x7f, palette) : undefined,
		],
	];
	for (const [key, value] of edges) if (value) Object.assign(xf.border, { [key]: value });
	return xf;
}

/** `FullColorExt` / `XFExtGradient`-free colour of an `XFEXT` property, or undefined when automatic. */
function fullColor(r: ByteReader, palette: readonly string[]): XlsColor | undefined {
	const type = r.u16();
	const tint = r.i16() / 32767;
	const value = r.u32();
	r.skip(8);
	const withTint = (color: XlsColor): XlsColor =>
		tint ? { ...color, tint: Math.round(tint * 1e6) / 1e6 } : color;
	if (type === 1) return withTint(paletteColor(value, palette));
	if (type === 2) {
		const rgb = [value & 0xff, (value >>> 8) & 0xff, (value >>> 16) & 0xff];
		return withTint({
			rgb: rgb.map((c) => c.toString(16).padStart(2, '0').toUpperCase()).join(''),
		});
	}
	if (type === 3) return withTint({ theme: value });
	return undefined;
}

/**
 * `XFEXT`: Excel 2007+ full colours for an XF. Applies foreground/background
 * fill, border and text colours over the palette approximations in place.
 */
export function applyXfExt(data: Uint8Array, xfs: XlsXf[], palette: readonly string[]): void {
	const r = new ByteReader(data, 14);
	const xf = xfs[r.u16()];
	r.skip(2);
	const count = r.u16();
	if (!xf) return;
	for (let i = 0; i < count && r.remaining >= 4; i++) {
		const start = r.pos;
		const type = r.u16();
		const size = r.u16();
		if (size < 4) break;
		const colorKey: Readonly<Record<number, 'top' | 'bottom' | 'left' | 'right' | 'diagonal'>> = {
			7: 'top',
			8: 'bottom',
			9: 'left',
			10: 'right',
			11: 'diagonal',
		};
		if (type === 4 || type === 5 || type === 13 || colorKey[type]) {
			const color = fullColor(r, palette);
			// A theme colour keeps the XF's palette approximation (index/rgb) as a fallback.
			const merge = (previous: XlsColor): XlsColor =>
				color?.theme !== undefined && previous.rgb
					? { index: previous.index ?? 0, rgb: previous.rgb, ...color }
					: (color ?? previous);
			if (color) {
				if (type === 4) xf.fill.fg = merge(xf.fill.fg);
				else if (type === 5) xf.fill.bg = merge(xf.fill.bg);
				else if (type === 13) xf.fontColor = color;
				else {
					const key = colorKey[type];
					const target = key ? xf.border[key] : undefined;
					if (target) target.color = merge(target.color);
				}
			}
		}
		r.pos = start + size;
	}
}
