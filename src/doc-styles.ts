/**
 * Stylesheet (STSH, [MS-DOC] 2.9.271): style definitions with names, type,
 * inheritance links and their own paragraph and character formatting.
 * Cascading is in `doc-style-resolve.ts`.
 *
 * @module doc-styles
 */
import { applyCharSprms } from './doc-char-props.js';
import type { DocCharProps, DocFontRefs } from './doc-char-props.js';
import { applyParaSprms } from './doc-para-props.js';
import type { DocParaProps } from './doc-para-props.js';
import { parseGrpprl } from './doc-sprm.js';
import type { DocRawSprm, SprmBudget } from './doc-sprm.js';

export type DocStyleType = 'paragraph' | 'character' | 'table' | 'list';

/** One style. `istd` is its index in the stylesheet; references are istd values too. */
export interface DocStyle {
	istd: number;
	/** Standard style identifier (0 Normal, 1..9 Heading n, ...); 0xFFE user defined, 0xFFF none. */
	sti: number;
	name: string;
	type: DocStyleType;
	/** istd of the style this one is based on. */
	basedOn?: number;
	/** istd of the style applied to the paragraph after this one. */
	next?: number;
	/** istd of the linked paragraph or character style. */
	link?: number;
	/** Built in (sti below 0xFFE) rather than user defined. */
	builtIn: boolean;
	/** The style's own paragraph formatting (not resolved through basedOn). */
	paragraph: DocParaProps;
	/** The style's own character formatting (not resolved through basedOn). */
	character: DocCharProps;
	/** Raw sprms behind {@link paragraph} and {@link character}, kept for lossless consumers. */
	paragraphSprms: readonly DocRawSprm[];
	characterSprms: readonly DocRawSprm[];
}

/** Parsed stylesheet. `styles` is sparse and indexed by istd. */
export interface DocStylesheet {
	styles: (DocStyle | undefined)[];
	/** Default fonts from the Stshi header (ftc indexes). */
	defaultFonts: DocFontRefs;
	/** Highest style identifier Word knew when it saved the file. */
	stiMaxWhenSaved: number;
}

/**
 * Names of the built-in styles by standard style identifier, used when a file stores no name.
 * Entries were checked against styles written by Word 16 (Title 62, Default Paragraph Font 65,
 * Subtitle 74, Hyperlink 85, the note and comment styles); the Followed Hyperlink, Strong and
 * Emphasis entries follow Hyperlink in sequence.
 */
const BUILTIN_NAMES: Record<number, string> = {
	0: 'Normal', 1: 'Heading 1', 2: 'Heading 2', 3: 'Heading 3', 4: 'Heading 4', 5: 'Heading 5',
	6: 'Heading 6', 7: 'Heading 7', 8: 'Heading 8', 9: 'Heading 9', 19: 'TOC 1', 29: 'Footnote Text',
	30: 'Comment Text', 31: 'Header', 32: 'Footer', 34: 'Caption', 38: 'Footnote Reference',
	39: 'Comment Reference', 41: 'Page Number', 42: 'Endnote Reference', 43: 'Endnote Text',
	62: 'Title', 65: 'Default Paragraph Font', 74: 'Subtitle', 85: 'Hyperlink',
	86: 'FollowedHyperlink', 87: 'Strong', 88: 'Emphasis', 105: 'Table Normal',
};

/** The built-in name for a standard style identifier, or undefined when none is known. */
export function builtInStyleName(sti: number): string | undefined {
	return BUILTIN_NAMES[sti];
}

const SGC: Record<number, DocStyleType> = { 1: 'paragraph', 2: 'character', 3: 'table', 4: 'list' };

/** Processing budget for the stylesheet. */
export interface StyleLimits {
	maxStyles: number;
	budget: SprmBudget;
}

/** Parse the STSH at `(fc, lcb)`; malformed individual styles are skipped. */
export function readDocStylesheet(table: Uint8Array, at: { fc: number; lcb: number }, limits: StyleLimits): DocStylesheet {
	const empty: DocStylesheet = { styles: [], defaultFonts: {}, stiMaxWhenSaved: 0 };
	if (at.lcb < 20 || at.fc > table.length - at.lcb) return empty;
	const v = new DataView(table.buffer, table.byteOffset, table.byteLength);
	const end = at.fc + at.lcb;
	const cbStshi = v.getUint16(at.fc, true);
	const head = at.fc + 2;
	if (cbStshi < 18 || head + cbStshi > end) return empty;
	const cstd = Math.min(v.getUint16(head, true), limits.maxStyles);
	const cbBase = v.getUint16(head + 2, true);
	const sheet: DocStylesheet = {
		styles: [],
		stiMaxWhenSaved: v.getUint16(head + 6, true),
		defaultFonts: {
			ascii: v.getUint16(head + 12, true),
			eastAsia: v.getUint16(head + 14, true),
			hAnsi: v.getUint16(head + 16, true),
			...(cbStshi >= 20 ? { cs: v.getUint16(head + 18, true) } : {}),
		},
	};
	let pos = head + cbStshi;
	for (let istd = 0; istd < cstd && pos + 2 <= end; istd++) {
		const cb = v.getUint16(pos, true);
		pos += 2;
		if (cb === 0) continue;
		if (pos + cb > end) break;
		const style = readStd(table.subarray(pos, pos + cb), istd, cbBase, limits.budget);
		if (style) sheet.styles[istd] = style;
		pos += cb;
	}
	return sheet;
}

function readStd(std: Uint8Array, istd: number, cbBase: number, budget: SprmBudget): DocStyle | undefined {
	if (cbBase < 10 || std.length < cbBase + 2) return undefined;
	const v = new DataView(std.buffer, std.byteOffset, std.byteLength);
	const w0 = v.getUint16(0, true), w1 = v.getUint16(2, true), w2 = v.getUint16(4, true);
	const sti = w0 & 0xfff;
	const type = SGC[w1 & 0xf];
	if (!type) return undefined;
	const cch = v.getUint16(cbBase, true);
	let at = cbBase + 2;
	if (at + cch * 2 + 2 > std.length) return undefined;
	let name = '';
	for (let i = 0; i < cch; i++) name += String.fromCharCode(v.getUint16(at + i * 2, true));
	at += cch * 2 + 2;
	if (at & 1) at++;
	const basedOn = w1 >> 4, next = w2 >> 4;
	const style: DocStyle = {
		istd, sti, type, builtIn: sti < 0xffe, name: name || builtInStyleName(sti) || '',
		paragraph: {}, character: {}, paragraphSprms: [], characterSprms: [],
	};
	if (basedOn !== 0xfff) style.basedOn = basedOn;
	if (next !== 0xfff && type === 'paragraph') style.next = next;
	if (cbBase >= 12) {
		const link = v.getUint16(10, true) & 0xfff;
		if (link !== 0 && link !== 0xfff) style.link = link;
	}
	const upx: Uint8Array[] = [];
	for (let i = 0, n = w2 & 0xf; i < n && at + 2 <= std.length; i++) {
		const cbUpx = v.getUint16(at, true);
		at += 2;
		if (at + cbUpx > std.length) break;
		upx.push(std.subarray(at, at + cbUpx));
		at += cbUpx + (cbUpx & 1);
	}
	applyUpx(style, upx, budget);
	return style;
}

/** UPX layout by style kind: paragraph [papx, chpx], character [chpx], table [tapx, papx, chpx], list [papx]. */
function applyUpx(style: DocStyle, upx: Uint8Array[], budget: SprmBudget): void {
	const papxAt = style.type === 'paragraph' ? 0 : style.type === 'table' ? 1 : style.type === 'list' ? 0 : -1;
	const chpxAt = style.type === 'paragraph' ? 1 : style.type === 'character' ? 0 : style.type === 'table' ? 2 : -1;
	const papx = papxAt >= 0 ? upx[papxAt] : undefined;
	if (papx && papx.length >= 2) {
		const sprms = parseGrpprl(papx.subarray(2), budget);
		style.paragraphSprms = sprms;
		applyParaSprms(style.paragraph, sprms);
	}
	const chpx = chpxAt >= 0 ? upx[chpxAt] : undefined;
	if (chpx) {
		const sprms = parseGrpprl(chpx, budget);
		style.characterSprms = sprms;
		applyCharSprms(style.character, sprms);
	}
}
