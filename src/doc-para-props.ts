/**
 * Paragraph formatting: decode PAPX/UpxPapx sprms ([MS-DOC] 2.6.2) onto a
 * typed property bag, including borders, shading and tab stops.
 *
 * @module doc-para-props
 */
import { docColorRef, docIcoToHex, i16, modernLast, sprmGroup, u16, u8 } from './doc-sprm.js';
import type { DocRawSprm } from './doc-sprm.js';

export type DocAlignment = 'left' | 'center' | 'right' | 'justify' | 'distribute';

/** One paragraph border edge. Widths and spacing are in points. */
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
}

export interface DocShading {
	/** Raw ipat (0 clear, 1 solid, 2.. percentages and hatch patterns). */
	pattern: number;
	foreground: string | undefined;
	background: string | undefined;
	/** Effective fill: the foreground for solid shading, otherwise the background. */
	fill: string | undefined;
}

export interface DocTabStop {
	positionTwips: number;
	alignment: 'left' | 'center' | 'right' | 'decimal' | 'bar';
	leader: 'none' | 'dot' | 'hyphen' | 'underscore' | 'heavy' | 'middleDot';
}

export interface DocLineSpacing {
	/** Auto: 240ths of a line; atLeast/exact: twips. */
	value: number;
	rule: 'auto' | 'atLeast' | 'exact';
}

export interface DocBorders {
	top?: DocBorder;
	left?: DocBorder;
	bottom?: DocBorder;
	right?: DocBorder;
	between?: DocBorder;
	bar?: DocBorder;
}

/** Typed paragraph properties. Absent fields are inherited or unspecified. */
export interface DocParaProps {
	/** Paragraph style (istd). */
	styleId?: number;
	alignment?: DocAlignment;
	indentLeftTwips?: number;
	indentRightTwips?: number;
	/** First-line indent in twips; negative values are a hanging indent. */
	firstLineTwips?: number;
	spaceBeforeTwips?: number;
	spaceAfterTwips?: number;
	spaceBeforeAuto?: boolean;
	spaceAfterAuto?: boolean;
	lineSpacing?: DocLineSpacing;
	keepTogether?: boolean;
	keepWithNext?: boolean;
	pageBreakBefore?: boolean;
	widowControl?: boolean;
	noLineNumbers?: boolean;
	noAutoHyphen?: boolean;
	contextualSpacing?: boolean;
	mirrorIndents?: boolean;
	bidi?: boolean;
	borders?: DocBorders;
	shading?: DocShading;
	tabs?: DocTabStop[];
	/** 0..8 for heading levels; absent for body text. */
	outlineLevel?: number;
	/** List instance (1-based index into the list overrides); 0 or absent means no list. */
	ilfo?: number;
	/** List level 0..8. */
	ilvl?: number;
	/** Paragraph is in a table cell. */
	inTable?: boolean;
	/** Paragraph mark is a table row end mark. */
	tableRowEnd?: boolean;
	/** Inner (nested) table cell or row end. */
	innerTable?: boolean;
	innerTableRowEnd?: boolean;
}

const JC: Record<number, DocAlignment> = { 0: 'left', 1: 'center', 2: 'right', 3: 'justify', 4: 'distribute' };

const BRC_NAMES: Record<number, string> = {
	0: 'none', 1: 'single', 2: 'thick', 3: 'double', 5: 'hairline', 6: 'dotted', 7: 'dashed',
	8: 'dotDash', 9: 'dotDotDash', 10: 'triple', 11: 'thinThickSmallGap', 12: 'thickThinSmallGap',
	13: 'thinThickThinSmallGap', 14: 'thinThickMediumGap', 15: 'thickThinMediumGap',
	16: 'thinThickThinMediumGap', 17: 'thinThickLargeGap', 18: 'thickThinLargeGap',
	19: 'thinThickThinLargeGap', 20: 'wave', 21: 'doubleWave', 22: 'dashSmallGap', 23: 'dashDotStroked',
	24: 'threeDEmboss', 25: 'threeDEngrave', 26: 'outset', 27: 'inset',
};

function border(type: number, width8: number, space: number, color: string | undefined, shadow: boolean): DocBorder {
	return { style: BRC_NAMES[type] ?? `type${type}`, type, widthPt: width8 / 8, spacePt: space, color, shadow };
}

/** Brc80 (4 bytes): width, type, ico, space/shadow bits. All-ones means no border. */
function brc80(o: Uint8Array): DocBorder | undefined {
	if (o.length < 4 || (o[0] === 0xff && o[1] === 0xff)) return undefined;
	return border(o[1]!, o[0]!, o[3]! & 0x1f, docIcoToHex(o[2]!), (o[3]! & 0x20) !== 0);
}

/** Brc (7 or 8 bytes): COLORREF, width, type, space/shadow bits. */
function brcModern(o: Uint8Array): DocBorder | undefined {
	if (o.length < 7 || (o[4] === 0xff && o[5] === 0xff)) return undefined;
	return border(o[5]!, o[4]!, o[6]! & 0x1f, docColorRef(o, 0), (o[6]! & 0x20) !== 0);
}

const BORDER_SLOTS: Record<number, [keyof DocBorders, boolean]> = {
	0x6424: ['top', false], 0x6425: ['left', false], 0x6426: ['bottom', false], 0x6427: ['right', false],
	0x6428: ['between', false], 0x6629: ['bar', false],
	0xc64e: ['top', true], 0xc64f: ['left', true], 0xc650: ['bottom', true], 0xc651: ['right', true],
	0xc652: ['between', true], 0xc653: ['bar', true],
};

function shading(fore: string | undefined, back: string | undefined, pattern: number): DocShading {
	return { pattern, foreground: fore, background: back, fill: pattern === 1 ? fore : back };
}

/** Shd80: icoFore (5 bits), icoBack (5 bits), ipat (6 bits). */
function shd80(o: Uint8Array): DocShading | undefined {
	const v = u16(o);
	if (v === undefined) return undefined;
	return shading(docIcoToHex(v & 0x1f), docIcoToHex((v >> 5) & 0x1f), v >> 10);
}

/** Shd: cvFore, cvBack, ipat. */
function shdModern(o: Uint8Array): DocShading | undefined {
	if (o.length < 10) return undefined;
	return shading(docColorRef(o, 0), docColorRef(o, 4), u16(o, 8)!);
}

const TAB_ALIGN = ['left', 'center', 'right', 'decimal', 'bar'] as const;
const TAB_LEADER = ['none', 'dot', 'hyphen', 'underscore', 'heavy', 'middleDot'] as const;

/** Merge a sprmPChgTabs/sprmPChgTabsPapx operand into the tab list (deletes, then adds). */
function applyTabs(p: DocParaProps, o: Uint8Array, withClose: boolean): void {
	let at = 0;
	const delCount = u8(o, at++) ?? 0;
	const deleted: number[] = [];
	for (let i = 0; i < delCount; i++) deleted.push(i16(o, at + i * 2) ?? 0);
	at += delCount * 2 + (withClose ? delCount * 2 : 0);
	const addCount = u8(o, at++) ?? 0;
	const tabs = (p.tabs ?? []).filter((t) => !deleted.includes(t.positionTwips));
	for (let i = 0; i < addCount; i++) {
		const pos = i16(o, at + i * 2);
		const tbd = u8(o, at + addCount * 2 + i);
		if (pos === undefined || tbd === undefined) break;
		const alignment = TAB_ALIGN[tbd & 7];
		if (!alignment) continue;
		const next: DocTabStop = { positionTwips: pos, alignment, leader: TAB_LEADER[(tbd >> 3) & 7] ?? 'none' };
		const same = tabs.findIndex((t) => t.positionTwips === pos);
		if (same >= 0) tabs[same] = next;
		else tabs.push(next);
	}
	tabs.sort((a, b) => a.positionTwips - b.positionTwips);
	p.tabs = tabs;
}

type FlagKey = 'keepTogether' | 'keepWithNext' | 'pageBreakBefore' | 'widowControl' | 'noLineNumbers'
	| 'noAutoHyphen' | 'contextualSpacing' | 'mirrorIndents' | 'bidi' | 'inTable' | 'tableRowEnd'
	| 'innerTable' | 'innerTableRowEnd' | 'spaceBeforeAuto' | 'spaceAfterAuto';

const FLAGS: Record<number, FlagKey> = {
	0x2405: 'keepTogether', 0x2406: 'keepWithNext', 0x2407: 'pageBreakBefore', 0x2431: 'widowControl',
	0x240c: 'noLineNumbers', 0x242a: 'noAutoHyphen', 0x246d: 'contextualSpacing', 0x2470: 'mirrorIndents',
	0x2441: 'bidi', 0x2416: 'inTable', 0x2417: 'tableRowEnd', 0x244b: 'innerTable', 0x244c: 'innerTableRowEnd',
	0x245b: 'spaceBeforeAuto', 0x245c: 'spaceAfterAuto',
};

/** Apply one paragraph sprm to `p`. Unknown or non-paragraph sprms are ignored. */
export function applyParaSprm(p: DocParaProps, s: DocRawSprm): void {
	const o = s.operand;
	const flagKey = FLAGS[s.opcode];
	if (flagKey) {
		const v = u8(o);
		if (v !== undefined) p[flagKey] = (v & 1) === 1;
		return;
	}
	const slot = BORDER_SLOTS[s.opcode];
	if (slot) {
		const b = slot[1] ? brcModern(o) : brc80(o);
		const borders = (p.borders ??= {});
		if (b) borders[slot[0]] = b;
		else delete borders[slot[0]];
		return;
	}
	switch (s.opcode) {
		case 0x4600: p.styleId = u16(o); break;
		case 0x2403: case 0x2461: { const a = JC[u8(o) ?? -1]; if (a) p.alignment = a; break; }
		case 0x840e: case 0x845d: p.indentRightTwips = i16(o); break;
		case 0x840f: case 0x845e: p.indentLeftTwips = i16(o); break;
		case 0x8411: case 0x8460: p.firstLineTwips = i16(o); break;
		case 0xa413: p.spaceBeforeTwips = u16(o); break;
		case 0xa414: p.spaceAfterTwips = u16(o); break;
		case 0x6412: {
			const line = i16(o), mult = u16(o, 2);
			if (line === undefined || mult === undefined) break;
			p.lineSpacing = mult === 1 ? { value: line, rule: 'auto' }
				: line < 0 ? { value: -line, rule: 'exact' } : { value: line, rule: 'atLeast' };
			break;
		}
		case 0x442d: { const v = shd80(o); if (v) p.shading = v; break; }
		case 0xc64d: { const v = shdModern(o); if (v) p.shading = v; break; }
		case 0x2640: {
			const v = u8(o);
			if (v !== undefined) {
				if (v <= 8) p.outlineLevel = v;
				else delete p.outlineLevel;
			}
			break;
		}
		case 0x260a: p.ilvl = u8(o); break;
		case 0x460b: p.ilfo = i16(o); break;
		case 0xc60d: applyTabs(p, o, false); break;
		case 0xc615: applyTabs(p, o, true); break;
		default: break;
	}
}

/** Apply a paragraph grpprl (legacy duplicates first), skipping non-paragraph groups. */
export function applyParaSprms(p: DocParaProps, sprms: readonly DocRawSprm[]): void {
	for (const s of modernLast(sprms)) if (sprmGroup(s.opcode) === 1) applyParaSprm(p, s);
}

/** Copy a property bag (borders and tabs are copied so styles are never mutated). */
export function cloneParaProps(p: DocParaProps): DocParaProps {
	const copy: DocParaProps = { ...p };
	if (p.borders) copy.borders = { ...p.borders };
	if (p.tabs) copy.tabs = p.tabs.map((t) => ({ ...t }));
	return copy;
}
