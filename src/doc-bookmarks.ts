/** Bookmarks: SttbfBkmk names joined with PlcfBkf starts and PlcfBkl ends
 * ([MS-DOC] 2.8.34 PlcfBkf, 2.8.35 PlcfBkl, 2.9.272 Sttbf, 2.9.20 BKF). */
import { DocResourceLimitError } from './ole-document-doc-pieces.js';
import { FCLCB, tableRange } from './doc-objects-context.js';
import type { DocObjectsContext } from './doc-objects-context.js';
import type { DocBookmark } from './doc-objects-types.js';

/** Read a Sttbf of UTF-16 strings (fExtend 0xFFFF); undefined for the ANSI form or damage. */
export function readExtendedSttbf(bytes: Uint8Array, max: number): string[] | undefined {
	if (bytes.length < 6 || bytes[0] !== 0xff || bytes[1] !== 0xff) return undefined;
	const count = bytes[2]! | (bytes[3]! << 8), cbExtra = bytes[4]! | (bytes[5]! << 8);
	if (count > max) throw new DocResourceLimitError('String table count exceeds the processing limit');
	const out: string[] = [];
	let at = 6;
	for (let i = 0; i < count; i++) {
		if (at + 2 > bytes.length) return undefined;
		const cch = bytes[at]! | (bytes[at + 1]! << 8);
		at += 2;
		if (at + cch * 2 + cbExtra > bytes.length) return undefined;
		let s = '';
		for (let j = 0; j < cch; j++) s += String.fromCharCode(bytes[at + j * 2]! | (bytes[at + j * 2 + 1]! << 8));
		out.push(s);
		at += cch * 2 + cbExtra;
	}
	return out;
}

export function readDocBookmarks(ctx: DocObjectsContext): DocBookmark[] {
	const names = tableRange(ctx, ctx.pair(FCLCB.sttbfBkmk));
	const bkf = tableRange(ctx, ctx.pair(FCLCB.plcfBkf));
	const bkl = tableRange(ctx, ctx.pair(FCLCB.plcfBkl));
	if (!names || !bkf || !bkl) return [];
	const list = readExtendedSttbf(names, ctx.limits.maxRecords);
	if (!list) { ctx.warnings.push('Bookmark name table is unreadable'); return []; }
	if (bkf.length < 4 || (bkf.length - 4) % 8 !== 0 || bkl.length < 4 || (bkl.length - 4) % 4 !== 0) {
		ctx.warnings.push('Bookmark position tables have invalid lengths');
		return [];
	}
	const n = (bkf.length - 4) / 8, nl = (bkl.length - 4) / 4;
	if (n > ctx.limits.maxRecords) throw new DocResourceLimitError('Bookmark count exceeds the processing limit');
	if (list.length !== n) ctx.warnings.push('Bookmark name count differs from the position table');
	const vf = new DataView(bkf.buffer, bkf.byteOffset, bkf.byteLength), vl = new DataView(bkl.buffer, bkl.byteOffset, bkl.byteLength);
	const out: DocBookmark[] = [];
	for (let i = 0; i < Math.min(n, list.length); i++) {
		const ibkl = vf.getUint16((n + 1) * 4 + i * 4, true), bkc = vf.getUint16((n + 1) * 4 + i * 4 + 2, true);
		if (ibkl >= nl) { ctx.warnings.push(`Bookmark "${list[i]}" has an invalid end index`); continue; }
		const cpStart = vf.getInt32(i * 4, true), cpEnd = vl.getInt32(ibkl * 4, true);
		if (cpStart < 0 || cpEnd < cpStart || cpEnd > ctx.text.length) { ctx.warnings.push(`Bookmark "${list[i]}" has an invalid range`); continue; }
		out.push({
			name: list[i]!, cpStart, cpEnd, story: ctx.storyOf(cpStart),
			hidden: list[i]!.startsWith('_'),
			columns: (bkc & 0x8000) ? { first: bkc & 0x7f, limit: (bkc >> 8) & 0x1f } : undefined,
		});
	}
	return out;
}
