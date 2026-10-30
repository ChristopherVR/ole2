/**
 * Border (Brc80, Brc) and shading (Shd80, Shd) decoders shared by paragraph,
 * table and cell properties ([MS-DOC] 2.9.25 Brc80Operand, 2.9.24 Brc,
 * 2.9.260 Shd80, 2.9.259 Shd).
 *
 * @module doc-brc
 */
import { docColorRef, docIcoToHex, u16 } from './doc-sprm.js';

/** One border edge. Widths and spacing are in points. */
export interface DocBorder {
	/** OOXML-style line name (`single`, `double`, `dotted`, ...) or `type<n>` when unnamed. */
	style: string;
	/** Raw brcType. */
	type: number;
	widthPt: number;
	spacePt: number;
	/** `RRGGBB`; undefined means automatic. */
	color: string | undefined;
	shadow: boolean;
	/** Border is drawn as a frame (`fFrame`). */
	frame: boolean;
}

/** Shading: pattern and colours; `fill` is the effective fill (foreground for solid, else background). */
export interface DocShading {
	/** Raw ipat (0 clear, 1 solid, 2.. percentages and hatch patterns). */
	pattern: number;
	foreground: string | undefined;
	background: string | undefined;
	fill: string | undefined;
}

const BRC_NAMES: Record<number, string> = {
	0: 'none', 1: 'single', 2: 'thick', 3: 'double', 5: 'hairline', 6: 'dotted', 7: 'dashed',
	8: 'dotDash', 9: 'dotDotDash', 10: 'triple', 11: 'thinThickSmallGap', 12: 'thickThinSmallGap',
	13: 'thinThickThinSmallGap', 14: 'thinThickMediumGap', 15: 'thickThinMediumGap',
	16: 'thinThickThinMediumGap', 17: 'thinThickLargeGap', 18: 'thickThinLargeGap',
	19: 'thinThickThinLargeGap', 20: 'wave', 21: 'doubleWave', 22: 'dashSmallGap', 23: 'dashDotStroked',
	24: 'threeDEmboss', 25: 'threeDEngrave', 26: 'outset', 27: 'inset',
};

function border(type: number, width8: number, flags: number, color: string | undefined): DocBorder {
	return {
		style: BRC_NAMES[type] ?? `type${type}`, type, widthPt: width8 / 8, spacePt: flags & 0x1f, color,
		shadow: (flags & 0x20) !== 0, frame: (flags & 0x40) !== 0,
	};
}

/** Brc80 (4 bytes at `at`): width, type, ico, space/shadow bits. All-ones or a zero line means no border. */
export function readBrc80(o: Uint8Array, at = 0): DocBorder | undefined {
	if (at + 4 > o.length || o[at + 1] === 0xff || (o[at + 1] === 0 && o[at] === 0)) return undefined;
	return border(o[at + 1]!, o[at]!, o[at + 3]!, docIcoToHex(o[at + 2]!));
}

/** Brc (7 or 8 bytes at `at`): COLORREF, width, type, space/shadow bits. */
export function readBrc(o: Uint8Array, at = 0): DocBorder | undefined {
	if (at + 7 > o.length || o[at + 5] === 0xff || (o[at + 5] === 0 && o[at + 4] === 0)) return undefined;
	return border(o[at + 5]!, o[at + 4]!, o[at + 6]!, docColorRef(o, at));
}

export function makeShading(fore: string | undefined, back: string | undefined, pattern: number): DocShading {
	return { pattern, foreground: fore, background: back, fill: pattern === 1 ? fore : back };
}

/** Shd80 (2 bytes): icoFore (5 bits), icoBack (5 bits), ipat (6 bits). Zero means no shading. */
export function readShd80(o: Uint8Array, at = 0): DocShading | undefined {
	const v = u16(o, at);
	if (v === undefined || v === 0) return undefined;
	return makeShading(docIcoToHex(v & 0x1f), docIcoToHex((v >> 5) & 0x1f), v >> 10);
}

/** Shd (10 bytes): cvFore, cvBack, ipat. An ipat of 0xFFFF means none. */
export function readShd(o: Uint8Array, at = 0): DocShading | undefined {
	if (at + 10 > o.length) return undefined;
	const pattern = u16(o, at + 8)!;
	if (pattern === 0xffff) return undefined;
	return makeShading(docColorRef(o, at), docColorRef(o, at + 4), pattern);
}
