/**
 * Font table (SttbfFfn, [MS-DOC] 2.9.272) and the Dop document properties
 * ([MS-DOC] 2.7.1) that are cheap to read.
 *
 * @module doc-fonts
 */

/** One entry of the font table. */
export interface DocFont {
	/** Index (ftc) character sprms use to reference this font. */
	index: number;
	name: string;
	/** FFN family: 0 unknown, 1 roman, 2 swiss, 3 modern, 4 script, 5 decorative. */
	family: number;
	/** FFN pitch: 0 default, 1 fixed, 2 variable. */
	pitch: number;
	trueType: boolean;
	weight: number;
	charset: number;
}

/** Name offset inside an Ffn after the leading byte: ffid, weight, chs, alt, panose, fs. */
const NAME_AT = 39;

/** Parse the font table at `(fc, lcb)`; stops quietly at the first malformed entry. */
export function readDocFonts(table: Uint8Array, at: { fc: number; lcb: number }, maxFonts = 4096): DocFont[] {
	if (at.lcb < 4 || at.fc > table.length - at.lcb) return [];
	const v = new DataView(table.buffer, table.byteOffset, table.byteLength);
	const count = Math.min(v.getUint16(at.fc, true), maxFonts);
	const end = at.fc + at.lcb;
	let pos = at.fc + 4;
	const fonts: DocFont[] = [];
	for (let i = 0; i < count && pos < end; i++) {
		const cbM1 = table[pos]!;
		const body = pos + 1;
		const next = body + cbM1;
		if (next > end || cbM1 < NAME_AT + 2) break;
		let name = '';
		for (let c = body + NAME_AT; c + 2 <= next; c += 2) {
			const code = v.getUint16(c, true);
			if (code === 0) break;
			name += String.fromCharCode(code);
		}
		const ffid = table[body]!;
		fonts.push({
			index: i,
			name,
			family: (ffid >> 4) & 7,
			pitch: ffid & 3,
			trueType: (ffid & 4) !== 0,
			weight: v.getUint16(body + 1, true),
			charset: table[body + 3]!,
		});
		pos = next;
	}
	return fonts;
}

/** The few Dop settings that are cheap and well defined. */
export interface DocProperties {
	/** Default tab stop in twips (dxaTab). */
	defaultTabStopTwips?: number;
	/** Different odd and even headers and footers. */
	facingPages?: boolean;
}

/** Read the basics of the Dop at `(fc, lcb)`; an absent or short Dop yields an empty result. */
export function readDocProperties(table: Uint8Array, at: { fc: number; lcb: number }): DocProperties {
	if (at.lcb < 12 || at.fc > table.length - at.lcb) return {};
	const v = new DataView(table.buffer, table.byteOffset, table.byteLength);
	const tab = v.getUint16(at.fc + 10, true);
	return {
		...(tab > 0 ? { defaultTabStopTwips: tab } : {}),
		facingPages: (table[at.fc]! & 1) === 1,
	};
}
