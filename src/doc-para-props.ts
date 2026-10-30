/**
 * Paragraph formatting: decode PAPX/UpxPapx sprms ([MS-DOC] 2.6.2) onto a
 * typed property bag, including borders, shading and tab stops.
 *
 * @module doc-para-props
 */
import { readBrc, readBrc80, readShd, readShd80 } from './doc-brc.js';
import type { DocBorder, DocShading } from './doc-brc.js';
import { i16, i32, modernLast, sprmGroup, u16, u8 } from './doc-sprm.js';
import type { DocRawSprm } from './doc-sprm.js';

export type DocAlignment = 'left' | 'center' | 'right' | 'justify' | 'distribute';


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
	/** Table nesting depth (sprmPItap); 0 or absent outside tables. */
	itap?: number;
}

const JC: Record<number, DocAlignment> = { 0: 'left', 1: 'center', 2: 'right', 3: 'justify', 4: 'distribute' };

const BORDER_SLOTS: Record<number, [keyof DocBorders, boolean]> = {
	0x6424: ['top', false], 0x6425: ['left', false], 0x6426: ['bottom', false], 0x6427: ['right', false],
	0x6428: ['between', false], 0x6629: ['bar', false],
	0xc64e: ['top', true], 0xc64f: ['left', true], 0xc650: ['bottom', true], 0xc651: ['right', true],
	0xc652: ['between', true], 0xc653: ['bar', true],
};

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
		const b = slot[1] ? readBrc(o) : readBrc80(o);
		const borders = (p.borders ??= {});
		if (b) borders[slot[0]] = b;
		else delete borders[slot[0]];
		return;
	}
	switch (s.opcode) {
		case 0x4600: p.styleId = u16(o); break;
		case 0x6649: p.itap = i32(o); break;
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
		case 0x442d: { const v = readShd80(o); if (v) p.shading = v; break; }
		case 0xc64d: { const v = readShd(o); if (v) p.shading = v; break; }
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
