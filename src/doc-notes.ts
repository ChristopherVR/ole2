/**
 * Footnotes, endnotes, comments and text boxes as story text ranges
 * ([MS-DOC] 2.8.? PlcffndRef/PlcffndTxt, PlcfandRef/PlcfandTxt,
 * PlcfendRef/PlcfendTxt, PlcftxbxTxt, PlcSpaMom/PlcSpaHdr, GrpXstAtnOwners).
 *
 * @module doc-notes
 */
import type { DocStoryId, DocStoryRange, StoryContext } from './doc-story.js';

export interface DocNote {
	readonly index: number;
	/** Main-story CP of the reference mark. */
	readonly refCp: number;
	/** True when Word numbers the note itself (the reference character is 0x02). */
	readonly autoNumbered: boolean;
	/** The note's text in the footnote or endnote story. */
	readonly range: DocStoryRange;
}

export interface DocComment {
	readonly index: number;
	/** Main-story CP of the comment reference mark. */
	readonly refCp: number;
	readonly author: string | undefined;
	readonly initials: string | undefined;
	/** `YYYY-MM-DDTHH:MM:00` local time when the file records a date. */
	readonly date: string | undefined;
	readonly range: DocStoryRange;
}

export interface DocTextBox {
	readonly index: number;
	/** `textboxes` for main-story shapes, `headerTextboxes` for shapes anchored in headers and footers. */
	readonly story: 'textboxes' | 'headerTextboxes';
	/** Drawing shape id (matches the shape in the drawing stream), when known. */
	readonly shapeId: number | undefined;
	/** CP of the shape anchor in the main story (or header story for header text boxes). */
	readonly anchorCp: number | undefined;
	/** Shape rectangle from the SPA in twips, relative to its anchoring frame. */
	readonly rectTwips: { left: number; top: number; right: number; bottom: number } | undefined;
	readonly range: DocStoryRange;
}

/** Footnotes (story `footnotes`) or endnotes (story `endnotes`). */
export function readNotes(ctx: StoryContext, kind: 'footnotes' | 'endnotes'): DocNote[] {
	const [refIdx, txtIdx] = kind === 'footnotes' ? [2, 3] : [46, 47];
	const refs = ctx.plc(refIdx, 2);
	const txt = ctx.plc(txtIdx, 0);
	if (!refs || !txt) return [];
	const tv = new DataView(ctx.table.buffer, ctx.table.byteOffset, ctx.table.byteLength);
	const out: DocNote[] = [];
	for (let i = 0; i < refs.count; i++) {
		if (i + 1 >= txt.cps.length) {
			ctx.warnings.push(`${kind} reference ${i} has no text`);
			break;
		}
		out.push({
			index: i,
			refCp: refs.cps[i]!,
			autoNumbered: tv.getInt16(refs.dataAt(i), true) !== 0,
			range: ctx.range(kind, txt.cps[i]!, txt.cps[i + 1]!),
		});
	}
	return out;
}

function readOwners(ctx: StoryContext): string[] {
	const at = ctx.pair(36);
	const owners: string[] = [];
	if (at.lcb === 0 || at.fc > ctx.table.length - at.lcb) return owners;
	const tv = new DataView(ctx.table.buffer, ctx.table.byteOffset, ctx.table.byteLength);
	const end = at.fc + at.lcb;
	for (let off = at.fc; off + 2 <= end && owners.length < ctx.limits.maxEntries; ) {
		const cch = tv.getUint16(off, true);
		if (off + 2 + cch * 2 > end) break;
		let s = '';
		for (let i = 0; i < cch; i++) s += String.fromCharCode(tv.getUint16(off + 2 + i * 2, true));
		owners.push(s);
		off += 2 + cch * 2;
	}
	return owners;
}

/** DTTM to an ISO-like local timestamp, or undefined for zero or invalid values. */
export function decodeDttm(v: number): string | undefined {
	if (v === 0) return undefined;
	const mint = v & 0x3f;
	const hr = (v >>> 6) & 0x1f;
	const dom = (v >>> 11) & 0x1f;
	const mon = (v >>> 16) & 0xf;
	const yr = (v >>> 20) & 0x1ff;
	if (mon < 1 || mon > 12 || dom < 1 || hr > 23 || mint > 59) return undefined;
	const p = (n: number, w = 2) => String(n).padStart(w, '0');
	return `${p(1900 + yr, 4)}-${p(mon)}-${p(dom)}T${p(hr)}:${p(mint)}:00`;
}

export function readComments(ctx: StoryContext): DocComment[] {
	const refs = ctx.plc(4, 30);
	const txt = ctx.plc(5, 0);
	if (!refs || !txt) return [];
	const owners = readOwners(ctx);
	const tv = new DataView(ctx.table.buffer, ctx.table.byteOffset, ctx.table.byteLength);
	const extra = ctx.pair(112);
	const stride = refs.count > 0 ? extra.lcb / refs.count : 0;
	const extraOk = Number.isInteger(stride) && stride >= 4 && extra.fc <= ctx.table.length - extra.lcb;
	const out: DocComment[] = [];
	for (let i = 0; i < refs.count; i++) {
		if (i + 1 >= txt.cps.length) {
			ctx.warnings.push(`comment ${i} has no text`);
			break;
		}
		const at = refs.dataAt(i);
		const cch = Math.min(tv.getUint16(at, true), 9);
		let initials = '';
		for (let c = 0; c < cch; c++) initials += String.fromCharCode(tv.getUint16(at + 2 + c * 2, true));
		const ibst = tv.getUint16(at + 20, true);
		out.push({
			index: i,
			refCp: refs.cps[i]!,
			author: owners[ibst],
			initials: initials || undefined,
			date: extraOk ? decodeDttm(tv.getUint32(extra.fc + i * stride, true)) : undefined,
			range: ctx.range('comments', txt.cps[i]!, txt.cps[i + 1]!),
		});
	}
	return out;
}

interface Spa {
	cp: number;
	lid: number;
	rect: { left: number; top: number; right: number; bottom: number };
}

function readSpas(ctx: StoryContext, index: number): Spa[] {
	const plc = ctx.plc(index, 26);
	if (!plc) return [];
	const tv = new DataView(ctx.table.buffer, ctx.table.byteOffset, ctx.table.byteLength);
	const out: Spa[] = [];
	for (let i = 0; i < plc.count; i++) {
		const at = plc.dataAt(i);
		out.push({
			cp: plc.cps[i]!,
			lid: tv.getInt32(at, true),
			rect: { left: tv.getInt32(at + 4, true), top: tv.getInt32(at + 8, true), right: tv.getInt32(at + 12, true), bottom: tv.getInt32(at + 16, true) },
		});
	}
	return out;
}

/** Text boxes of the main story and of headers/footers. */
export function readTextBoxes(ctx: StoryContext): DocTextBox[] {
	const out: DocTextBox[] = [];
	const groups: Array<{ story: 'textboxes' | 'headerTextboxes'; txt: number; spa: number }> = [
		{ story: 'textboxes', txt: 56, spa: 40 },
		{ story: 'headerTextboxes', txt: 58, spa: 41 },
	];
	const tv = new DataView(ctx.table.buffer, ctx.table.byteOffset, ctx.table.byteLength);
	for (const g of groups) {
		const plc = ctx.plc(g.txt, 22);
		if (!plc) continue;
		const spas = new Map(readSpas(ctx, g.spa).map((s) => [s.lid, s]));
		// The last PlcfTxbxTxt entry is a sentinel story, not a text box.
		for (let i = 0; i < plc.count - 1; i++) {
			const lid = tv.getInt32(plc.dataAt(i) + 14, true);
			const spa = spas.get(lid);
			out.push({
				index: out.length,
				story: g.story,
				shapeId: lid || undefined,
				anchorCp: spa?.cp,
				rectTwips: spa?.rect,
				range: ctx.range(g.story as DocStoryId, plc.cps[i]!, plc.cps[i + 1]!),
			});
		}
	}
	return out;
}
