/**
 * Lists: PlfLst/LSTF/LVL list definitions and PlfLfo/LFO list instances with
 * level overrides ([MS-DOC] 2.9.147, 2.9.150, 2.9.151, 2.9.153).
 *
 * @module doc-lists
 */
import { applyCharSprms } from './doc-char-props.js';
import type { DocCharProps } from './doc-char-props.js';
import { applyParaSprms } from './doc-para-props.js';
import type { DocParaProps } from './doc-para-props.js';
import { parseGrpprl } from './doc-sprm.js';
import type { SprmBudget } from './doc-sprm.js';

/** One list level (LVL). */
export interface DocListLevel {
	/** Level 0..8. */
	level: number;
	startAt: number;
	/** Raw number format code (nfc): 0 arabic, 1 upper roman, 2 lower roman, 3 upper letter, 4 lower letter, 23 bullet, 255 none. */
	numberFormat: number;
	/** OOXML-style name for common formats, else `nfc<n>`. */
	numberFormatName: string;
	/** Level text with `%1`..`%9` placeholders for the level counters; the bullet character for bullets. */
	text: string;
	isBullet: boolean;
	alignment: 'left' | 'center' | 'right';
	/** Follow the number with a tab, a space or nothing. */
	follow: 'tab' | 'space' | 'nothing';
	/** Show earlier levels in arabic numerals (legal numbering). */
	legal: boolean;
	/** Do not restart numbering after a higher level. */
	noRestart: boolean;
	/** Paragraph formatting the level applies (indents, tabs). */
	paragraph: DocParaProps;
	/** Number or bullet character formatting, including the bullet font. */
	character: DocCharProps;
}

/** A list definition (LSTF with its levels). */
export interface DocListDefinition {
	lsid: number;
	templateId: number;
	/** Styles linked to levels (istd), 0xFFF entries removed; index is the level. */
	styleIds: (number | undefined)[];
	/** Single level list (one LVL instead of nine). */
	simple: boolean;
	hybrid: boolean;
	levels: DocListLevel[];
}

export interface DocListLevelOverride {
	level: number;
	startAt?: number;
	/** Replacement level formatting when the override carries one. */
	definition?: DocListLevel;
}

/** A list instance (LFO): what `ilfo` of a paragraph selects. */
export interface DocListOverride {
	/** 1-based, matches `ilfo`. */
	ilfo: number;
	lsid: number;
	/** Index into {@link DocLists.definitions}, undefined when the lsid is unknown. */
	definitionIndex?: number;
	levels: DocListLevelOverride[];
}

export interface DocLists {
	definitions: DocListDefinition[];
	overrides: DocListOverride[];
}

const NFC: Record<number, string> = {
	0: 'decimal', 1: 'upperRoman', 2: 'lowerRoman', 3: 'upperLetter', 4: 'lowerLetter', 5: 'ordinal',
	6: 'cardinalText', 7: 'ordinalText', 22: 'decimalZero', 23: 'bullet', 255: 'none',
};

export interface ListLimits {
	maxLists: number;
	budget: SprmBudget;
	fontNames: (props: DocCharProps) => DocCharProps;
}

/** Parse one LVL at `pos`; returns undefined when it does not fit before `end`. */
function readLvl(b: Uint8Array, pos: number, end: number, level: number, lim: ListLimits): { lvl: DocListLevel; next: number } | undefined {
	if (pos + 28 > end) return undefined;
	const v = new DataView(b.buffer, b.byteOffset, b.byteLength);
	const nfc = b[pos + 4]!;
	const flags = b[pos + 5]!;
	const cbChpx = b[pos + 24]!, cbPapx = b[pos + 25]!;
	let at = pos + 28;
	if (at + cbPapx + cbChpx + 2 > end) return undefined;
	const paragraph: DocParaProps = {};
	const character: DocCharProps = {};
	applyParaSprms(paragraph, parseGrpprl(b.subarray(at, at + cbPapx), lim.budget));
	at += cbPapx;
	applyCharSprms(character, parseGrpprl(b.subarray(at, at + cbChpx), lim.budget));
	at += cbChpx;
	const cch = v.getUint16(at, true);
	at += 2;
	if (at + cch * 2 > end) return undefined;
	const codes: number[] = [];
	for (let i = 0; i < cch; i++) codes.push(v.getUint16(at + i * 2, true));
	at += cch * 2;
	let text = '';
	const placeholders = new Map<number, number>();
	for (let k = 0; k < 9; k++) {
		const p = b[pos + 6 + k]!;
		if (p > 0 && p <= cch) placeholders.set(p - 1, codes[p - 1]!);
	}
	for (let i = 0; i < cch; i++) {
		const ph = placeholders.get(i);
		text += ph !== undefined && ph < 9 ? `%${ph + 1}` : String.fromCharCode(codes[i]!);
	}
	const lvl: DocListLevel = {
		level,
		startAt: v.getInt32(pos, true),
		numberFormat: nfc,
		numberFormatName: NFC[nfc] ?? `nfc${nfc}`,
		text,
		isBullet: nfc === 23,
		alignment: (['left', 'center', 'right'] as const)[flags & 3] ?? 'left',
		follow: (['tab', 'space', 'nothing'] as const)[b[pos + 15]!] ?? 'tab',
		legal: (flags & 4) !== 0,
		noRestart: (flags & 8) !== 0,
		paragraph,
		character: lim.fontNames(character),
	};
	return { lvl, next: at };
}

/** Parse PlfLst and PlfLfo. Malformed tails are dropped rather than failing the whole document. */
export function readDocLists(
	table: Uint8Array,
	plfLst: { fc: number; lcb: number },
	plfLfo: { fc: number; lcb: number },
	lim: ListLimits,
): DocLists {
	// The LVL records follow the PlfLst but are not counted in its lcb; they end where PlfLfo starts.
	const lvlEnd = plfLfo.lcb > 0 && plfLfo.fc >= plfLst.fc + plfLst.lcb ? plfLfo.fc : table.length;
	const definitions = readDefinitions(table, plfLst, lvlEnd, lim);
	return { definitions, overrides: readOverrides(table, plfLfo, definitions, lim) };
}

function readDefinitions(table: Uint8Array, at: { fc: number; lcb: number }, lvlEnd: number, lim: ListLimits): DocListDefinition[] {
	if (at.lcb < 2 || at.fc > table.length - at.lcb) return [];
	const v = new DataView(table.buffer, table.byteOffset, table.byteLength);
	const end = Math.min(lvlEnd, table.length);
	const count = Math.min(v.getInt16(at.fc, true), lim.maxLists);
	if (count <= 0 || at.fc + 2 + count * 28 > at.fc + at.lcb) return [];
	const heads = at.fc + 2;
	let pos = at.fc + at.lcb;
	const out: DocListDefinition[] = [];
	for (let i = 0; i < count; i++) {
		const h = heads + i * 28;
		const flags = table[h + 26]!;
		const simple = (flags & 1) !== 0;
		const styleIds: (number | undefined)[] = [];
		for (let k = 0; k < 9; k++) {
			const s = v.getUint16(h + 8 + k * 2, true);
			styleIds.push(s === 0xfff ? undefined : s);
		}
		const levels: DocListLevel[] = [];
		for (let k = 0; k < (simple ? 1 : 9); k++) {
			const r = readLvl(table, pos, end, k, lim);
			if (!r) return out.concat(listOf(v, h, simple, styleIds, levels));
			levels.push(r.lvl);
			pos = r.next;
		}
		out.push(...listOf(v, h, simple, styleIds, levels));
	}
	return out;
}

function listOf(v: DataView, h: number, simple: boolean, styleIds: (number | undefined)[], levels: DocListLevel[]): DocListDefinition[] {
	if (levels.length === 0) return [];
	return [{ lsid: v.getInt32(h, true), templateId: v.getInt32(h + 4, true), styleIds, simple, hybrid: (v.getUint8(h + 26) & 0x10) !== 0, levels }];
}

function readOverrides(table: Uint8Array, at: { fc: number; lcb: number }, defs: DocListDefinition[], lim: ListLimits): DocListOverride[] {
	if (at.lcb < 4 || at.fc > table.length - at.lcb) return [];
	const v = new DataView(table.buffer, table.byteOffset, table.byteLength);
	const end = at.fc + at.lcb;
	const count = Math.min(v.getInt32(at.fc, true), lim.maxLists);
	if (count <= 0 || at.fc + 4 + count * 16 > end) return [];
	let pos = at.fc + 4 + count * 16;
	const out: DocListOverride[] = [];
	for (let i = 0; i < count; i++) {
		const h = at.fc + 4 + i * 16;
		const lsid = v.getInt32(h, true);
		const clfolvl = table[h + 12]!;
		const levels: DocListLevelOverride[] = [];
		pos += 4; // LFOData.cp
		for (let k = 0; k < clfolvl; k++) {
			if (pos + 8 > end) break;
			const flags = v.getUint32(pos + 4, true);
			const ov: DocListLevelOverride = { level: flags & 0xf };
			if (flags & 0x10) ov.startAt = v.getInt32(pos, true);
			pos += 8;
			if (flags & 0x20) {
				const r = readLvl(table, pos, end, ov.level, lim);
				if (!r) break;
				ov.definition = r.lvl;
				pos = r.next;
			}
			levels.push(ov);
		}
		const definitionIndex = defs.findIndex((d) => d.lsid === lsid);
		out.push({ ilfo: i + 1, lsid, ...(definitionIndex >= 0 ? { definitionIndex } : {}), levels });
	}
	return out;
}
