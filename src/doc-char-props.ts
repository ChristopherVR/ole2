/**
 * Character formatting: decode CHPX/UpxChpx sprms ([MS-DOC] 2.6.1) onto a
 * typed property bag. Appliers mutate a working object; callers cascade
 * defaults, paragraph style, character style and direct exceptions in order.
 *
 * @module doc-char-props
 */
import { docColorRef, docIcoToHex, i16, modernLast, sprmGroup, u16, u8 } from './doc-sprm.js';
import type { DocRawSprm } from './doc-sprm.js';

/** Underline kinds named after the OOXML `ST_Underline` values they map to. */
export type DocUnderlineKind =
	| 'none' | 'single' | 'words' | 'double' | 'dotted' | 'thick' | 'dash' | 'dotDash' | 'dotDotDash'
	| 'wave' | 'dottedHeavy' | 'dashHeavy' | 'dotDashHeavy' | 'dotDotDashHeavy' | 'waveHeavy'
	| 'dashLong' | 'dashLongHeavy' | 'waveDouble';

/** Font table indexes (ftc) per script class. */
export interface DocFontRefs {
	ascii?: number;
	eastAsia?: number;
	hAnsi?: number;
	cs?: number;
}

/** Resolved font names per script class. */
export interface DocFontNames {
	ascii?: string;
	eastAsia?: string;
	hAnsi?: string;
	cs?: string;
}

/** Typed character properties. Absent fields are inherited or unspecified, never false. */
export interface DocCharProps {
	/** Character style (istd) from sprmCIstd. */
	styleId?: number;
	bold?: boolean;
	italic?: boolean;
	underline?: DocUnderlineKind;
	/** Underline colour, `RRGGBB`. */
	underlineColor?: string;
	strike?: boolean;
	doubleStrike?: boolean;
	caps?: boolean;
	smallCaps?: boolean;
	hidden?: boolean;
	outline?: boolean;
	shadow?: boolean;
	emboss?: boolean;
	imprint?: boolean;
	verticalAlign?: 'superscript' | 'subscript';
	/** Font size in points (half-point precision). */
	sizePt?: number;
	sizeCsPt?: number;
	/** Text colour `RRGGBB`; undefined means automatic. */
	color?: string;
	/** Highlight colour name (`yellow`, `darkBlue`, ...). */
	highlight?: string;
	fontIds?: DocFontRefs;
	fonts?: DocFontNames;
	/** Expanded (+) or condensed (-) spacing in twips. */
	spacingTwips?: number;
	/** Raised (+) or lowered (-) baseline offset in points. */
	positionPt?: number;
	/** Horizontal scale in percent. */
	scalePercent?: number;
	/** Kerning threshold in points (0 disables). */
	kerningPt?: number;
	/** Complex-script bold and italic. */
	boldCs?: boolean;
	italicCs?: boolean;
	rtl?: boolean;
	noProof?: boolean;
	/** Run is a special character (picture, footnote reference, field mark, ...). */
	special?: boolean;
	/** Field result is hidden. */
	fieldVanish?: boolean;
}

const UNDERLINES: Record<number, DocUnderlineKind> = {
	0: 'none', 1: 'single', 2: 'words', 3: 'double', 4: 'dotted', 5: 'thick', 6: 'dash', 7: 'dotDash',
	8: 'dotDotDash', 9: 'wave', 0x14: 'dottedHeavy', 0x15: 'dashHeavy', 0x16: 'dotDashHeavy',
	0x17: 'dotDotDashHeavy', 0x18: 'waveHeavy', 0x19: 'dashLong', 0x1a: 'dashLongHeavy', 0x2b: 'waveDouble',
};

const HIGHLIGHTS = [
	undefined, 'black', 'blue', 'cyan', 'green', 'magenta', 'red', 'yellow', 'white', 'darkBlue',
	'darkCyan', 'darkGreen', 'darkMagenta', 'darkRed', 'darkYellow', 'darkGray', 'lightGray',
] as const;

type BoolKey = 'bold' | 'italic' | 'strike' | 'doubleStrike' | 'caps' | 'smallCaps' | 'hidden' | 'outline'
	| 'shadow' | 'emboss' | 'imprint' | 'boldCs' | 'italicCs' | 'rtl' | 'noProof' | 'special' | 'fieldVanish';

const TOGGLES: Record<number, BoolKey> = {
	0x0835: 'bold', 0x0836: 'italic', 0x0837: 'strike', 0x0838: 'outline', 0x0839: 'shadow',
	0x083a: 'smallCaps', 0x083b: 'caps', 0x083c: 'hidden', 0x2a53: 'doubleStrike', 0x0854: 'imprint',
	0x0858: 'emboss', 0x085c: 'boldCs', 0x085d: 'italicCs', 0x085a: 'rtl', 0x0875: 'noProof',
	0x0855: 'special', 0x0802: 'fieldVanish',
};

/** Toggle operand: 0 off, 1 on, 0x80 same as style (no change), 0x81 opposite of style. */
function applyToggle(p: DocCharProps, key: BoolKey, value: number): void {
	if (value === 0) p[key] = false;
	else if (value === 1) p[key] = true;
	else if (value === 0x81) p[key] = !p[key];
}

function setFont(p: DocCharProps, slot: keyof DocFontRefs, ftc: number): void {
	(p.fontIds ??= {})[slot] = ftc;
}

/** Apply one character sprm to `p`. Unknown or non-character sprms are ignored. */
export function applyCharSprm(p: DocCharProps, s: DocRawSprm): void {
	const o = s.operand;
	const toggle = TOGGLES[s.opcode];
	if (toggle) {
		const v = u8(o);
		if (v !== undefined) applyToggle(p, toggle, v);
		return;
	}
	switch (s.opcode) {
		case 0x4a30: p.styleId = u16(o); break;
		case 0x2a3e: {
			const kind = UNDERLINES[u8(o) ?? -1];
			if (kind) p.underline = kind;
			break;
		}
		case 0x6877: p.underlineColor = docColorRef(o); break;
		case 0x2a42: {
			const ico = u8(o) ?? 0;
			p.color = docIcoToHex(ico);
			break;
		}
		case 0x6870: p.color = docColorRef(o); break;
		case 0x2a0c: {
			const name = HIGHLIGHTS[u8(o) ?? 0];
			if (name) p.highlight = name;
			else delete p.highlight;
			break;
		}
		case 0x4a43: { const v = u16(o); if (v !== undefined) p.sizePt = v / 2; break; }
		case 0x4a61: { const v = u16(o); if (v !== undefined) p.sizeCsPt = v / 2; break; }
		case 0x2a48: {
			const v = u8(o);
			if (v === 1) p.verticalAlign = 'superscript';
			else if (v === 2) p.verticalAlign = 'subscript';
			else if (v === 0) delete p.verticalAlign;
			break;
		}
		case 0x4a4f: { const v = u16(o); if (v !== undefined) setFont(p, 'ascii', v); break; }
		case 0x4a50: { const v = u16(o); if (v !== undefined) setFont(p, 'eastAsia', v); break; }
		case 0x4a51: { const v = u16(o); if (v !== undefined) setFont(p, 'hAnsi', v); break; }
		case 0x4a5e: { const v = u16(o); if (v !== undefined) setFont(p, 'cs', v); break; }
		case 0x8840: p.spacingTwips = i16(o); break;
		case 0x4845: { const v = i16(o); if (v !== undefined) p.positionPt = v / 2; break; }
		case 0x4852: p.scalePercent = u16(o); break;
		case 0x484b: { const v = u16(o); if (v !== undefined) p.kerningPt = v / 2; break; }
		default: break;
	}
}

/** Apply a character grpprl (legacy duplicates first), skipping non-character groups. */
export function applyCharSprms(p: DocCharProps, sprms: readonly DocRawSprm[]): void {
	for (const s of modernLast(sprms)) if (sprmGroup(s.opcode) === 2) applyCharSprm(p, s);
}

/** Copy a property bag (one level deep for the font sub-objects). */
export function cloneCharProps(p: DocCharProps): DocCharProps {
	const copy: DocCharProps = { ...p };
	if (p.fontIds) copy.fontIds = { ...p.fontIds };
	if (p.fonts) copy.fonts = { ...p.fonts };
	return copy;
}
