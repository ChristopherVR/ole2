/**
 * Section properties from `PlcfSed` and each section's `Sepx`
 * ([MS-DOC] 2.8.25 Sed, 2.9.245 Sepx, 2.6.4 section sprms), plus the
 * document-wide facing-pages flag from the `Dop`.
 *
 * @module doc-sections
 */
import type { StoryContext } from './doc-story.js';
import { findSprm as lastSprm, i16, parseGrpprl, u16, u8 } from './doc-sprm.js';
import type { DocRawSprm } from './doc-sprm.js';

const operandU8 = (s: DocRawSprm | undefined) => (s ? u8(s.operand) : undefined);
const operandU16 = (s: DocRawSprm | undefined) => (s ? u16(s.operand) : undefined);
const operandI16 = (s: DocRawSprm | undefined) => (s ? i16(s.operand) : undefined);

export type DocSectionBreak = 'continuous' | 'newColumn' | 'nextPage' | 'evenPage' | 'oddPage';
export type DocPageNumberFormat = 'decimal' | 'upperRoman' | 'lowerRoman' | 'upperLetter' | 'lowerLetter' | 'other';

export interface DocSectionColumn {
	readonly widthTwips: number;
	/** Space after this column; undefined for the last one. */
	readonly spacingTwips: number | undefined;
}

/** One section. Lengths are twips (1/20 pt); absent sprms take the [MS-DOC] defaults. */
export interface DocSection {
	readonly index: number;
	/** Main-story CP range `[start, end)` the section covers. */
	readonly start: number;
	readonly end: number;
	readonly breakType: DocSectionBreak;
	readonly pageWidthTwips: number;
	readonly pageHeightTwips: number;
	readonly orientation: 'portrait' | 'landscape';
	readonly marginLeftTwips: number;
	readonly marginRightTwips: number;
	readonly marginTopTwips: number;
	readonly marginBottomTwips: number;
	readonly gutterTwips: number;
	readonly headerDistanceTwips: number;
	readonly footerDistanceTwips: number;
	readonly columnCount: number;
	readonly columnSpacingTwips: number;
	readonly columnsEvenlySpaced: boolean;
	readonly lineBetweenColumns: boolean;
	/** Explicit per-column widths when the columns are not evenly spaced. */
	readonly columns: readonly DocSectionColumn[];
	/** First page uses its own header/footer (sprmSFTitlePage). */
	readonly titlePage: boolean;
	readonly pageNumberFormat: DocPageNumberFormat;
	/** Number the section's first page starts from, when restart or a start value is set. */
	readonly pageNumberStart: number | undefined;
	readonly pageNumberRestart: boolean;
	/** Raw section sprms, for callers that decode more. */
	readonly sprms: readonly DocRawSprm[];
}

const BREAKS: readonly DocSectionBreak[] = ['continuous', 'newColumn', 'nextPage', 'evenPage', 'oddPage'];
const FORMATS: readonly DocPageNumberFormat[] = ['decimal', 'upperRoman', 'lowerRoman', 'upperLetter', 'lowerLetter'];
const PLCF_SED = 6;
const DOP = 31;

/** Whether the document uses distinct odd and even headers (Dop.fFacingPages). */
export function readFacingPages(ctx: StoryContext): boolean {
	const dop = ctx.pair(DOP);
	if (dop.lcb < 1 || dop.fc > ctx.table.length - 1) return false;
	return (ctx.table[dop.fc]! & 1) !== 0;
}

/** Read every section; returns an empty list when the document has no PlcfSed. */
export function readDocSections(ctx: StoryContext): DocSection[] {
	const plc = ctx.plc(PLCF_SED, 12);
	if (!plc) return [];
	const wv = new DataView(ctx.word.buffer, ctx.word.byteOffset, ctx.word.byteLength);
	const tv = new DataView(ctx.table.buffer, ctx.table.byteOffset, ctx.table.byteLength);
	const mainLength = ctx.extents.main.length;
	const out: DocSection[] = [];
	for (let i = 0; i < plc.count; i++) {
		const fcSepx = tv.getUint32(plc.dataAt(i) + 2, true);
		let sprms: readonly DocRawSprm[] = [];
		if (fcSepx !== 0xffffffff) {
			if (fcSepx + 2 > ctx.word.length) {
				ctx.warnings.push(`Section ${i} Sepx lies outside WordDocument`);
			} else {
				const cb = wv.getUint16(fcSepx, true);
				sprms = parseGrpprl(ctx.word.subarray(fcSepx + 2, Math.min(fcSepx + 2 + cb, ctx.word.length)), ctx.budget);
			}
		}
		out.push(decodeSection(i, plc.cps[i]!, Math.min(plc.cps[i + 1]!, mainLength), sprms));
	}
	return out;
}

function decodeSection(index: number, start: number, end: number, sprms: readonly DocRawSprm[]): DocSection {
	const num = (op: number, dflt: number, signed = false) => (signed ? operandI16(lastSprm(sprms, op)) : operandU16(lastSprm(sprms, op))) ?? dflt;
	const flag = (op: number) => (operandU8(lastSprm(sprms, op)) ?? 0) !== 0;
	const columnCount = (operandU16(lastSprm(sprms, 0x500b)) ?? 0) + 1;
	const spacing = num(0x900c, 720);
	const evenly = lastSprm(sprms, 0x3005) ? flag(0x3005) : true;
	const widths = new Map<number, number>();
	const gaps = new Map<number, number>();
	for (const s of sprms) {
		if ((s.opcode === 0xf203 || s.opcode === 0xf204) && s.operand.length >= 3) {
			(s.opcode === 0xf203 ? widths : gaps).set(s.operand[0]!, s.operand[1]! | (s.operand[2]! << 8));
		}
	}
	const columns: DocSectionColumn[] = [];
	if (!evenly) {
		for (let c = 0; c < columnCount; c++) {
			columns.push({ widthTwips: widths.get(c) ?? 0, spacingTwips: c < columnCount - 1 ? (gaps.get(c) ?? spacing) : undefined });
		}
	}
	const fmt = operandU8(lastSprm(sprms, 0x300e));
	const startRaw = lastSprm(sprms, 0x501c);
	const restart = flag(0x3011);
	return {
		index,
		start,
		end,
		breakType: BREAKS[operandU8(lastSprm(sprms, 0x3009)) ?? 2] ?? 'nextPage',
		pageWidthTwips: num(0xb01f, 12240),
		pageHeightTwips: num(0xb020, 15840),
		orientation: operandU8(lastSprm(sprms, 0x301d)) === 2 ? 'landscape' : 'portrait',
		marginLeftTwips: num(0xb021, 1800),
		marginRightTwips: num(0xb022, 1800),
		marginTopTwips: num(0x9023, 1440, true),
		marginBottomTwips: num(0x9024, 1440, true),
		gutterTwips: num(0xb025, 0),
		headerDistanceTwips: num(0xb017, 720),
		footerDistanceTwips: num(0xb018, 720),
		columnCount,
		columnSpacingTwips: spacing,
		columnsEvenlySpaced: evenly,
		lineBetweenColumns: flag(0x3019),
		columns,
		titlePage: flag(0x300a),
		pageNumberFormat: fmt === undefined ? 'decimal' : (FORMATS[fmt] ?? 'other'),
		pageNumberStart: startRaw ? operandU16(startRaw) : undefined,
		pageNumberRestart: restart,
		sprms,
	};
}
