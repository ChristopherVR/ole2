/**
 * Shared context for reading the non-table structure of a legacy `.doc`:
 * the story layout of the document-wide character position (CP) space
 * ([MS-DOC] 2.4.1, FibRgLw97), decoded text, and bounded PLC access.
 *
 * @module doc-story
 */
import { unwrapDocBytes } from './ole-document-doc-cfb.js';
import { readDocFib, readFcLcbAt, type ParsedDocFib } from './ole-document-doc-fib.js';
import { DocResourceLimitError, decodePiecesText, parseDocClx, type DocPiece } from './ole-document-doc-pieces.js';

/** The seven stories of a Word document, in CP order. */
export type DocStoryId = 'main' | 'footnotes' | 'headers' | 'comments' | 'endnotes' | 'textboxes' | 'headerTextboxes';

/** Application processing budgets; none is a format validity limit. */
export interface DocStructureLimits {
	/** UTF-16 positions across all stories (default 16,777,216). */
	maxCharacters?: number;
	/** Pieces in the piece table (default 65,536). */
	maxPieces?: number;
	/** Paragraphs scanned for table structure across all stories (default 1,000,000). */
	maxParagraphs?: number;
	/** Table cells produced (default 262,144). */
	maxCells?: number;
	/** Entries of any one PLC (sections, notes, comments, text boxes, header stories; default 65,536). */
	maxEntries?: number;
	/** Property modifiers decoded (default 1,048,576). */
	maxSprms?: number;
}

export const DEFAULT_DOC_STRUCTURE_LIMITS: Readonly<Required<DocStructureLimits>> = Object.freeze({
	maxCharacters: 16_777_216,
	maxPieces: 65_536,
	maxParagraphs: 1_000_000,
	maxCells: 262_144,
	maxEntries: 65_536,
	maxSprms: 1_048_576,
});

/** A character range in one story; positions are relative to the story start. */
export interface DocStoryRange {
	readonly story: DocStoryId;
	/** Story-relative start CP. */
	readonly start: number;
	/** Story-relative end CP (exclusive). */
	readonly end: number;
	/** Document-wide CP of `start` (equals `start` for the main story). */
	readonly globalStart: number;
	/** Raw characters of the range, including paragraph, cell and reference marks. */
	readonly text: string;
}

export interface StoryExtent {
	readonly start: number;
	readonly length: number;
}

export interface Plc {
	readonly cps: readonly number[];
	/** Number of data elements. */
	readonly count: number;
	/** Absolute table-stream offset of data element `i`. */
	dataAt(i: number): number;
}

export interface StoryContext {
	readonly word: Uint8Array;
	readonly table: Uint8Array;
	readonly fib: ParsedDocFib;
	readonly pieces: readonly DocPiece[];
	readonly text: string;
	readonly limits: Required<DocStructureLimits>;
	readonly extents: Readonly<Record<DocStoryId, StoryExtent>>;
	readonly warnings: string[];
	/** Raw FibRgFcLcb97/2000/2002 pair, or empty when the index is beyond this FIB's array. */
	pair(index: number): { fc: number; lcb: number };
	range(story: DocStoryId, start: number, end: number): DocStoryRange;
	/** Byte offset in WordDocument of CP `cp` (document-wide), or undefined if no piece covers it. */
	fcAt(cp: number): number | undefined;
	plc(index: number, dataSize: number): Plc | undefined;
}

const ORDER: readonly DocStoryId[] = ['main', 'footnotes', 'headers', 'comments', 'endnotes', 'textboxes', 'headerTextboxes'];
// FibRgLw97 element index of each story length (ccpText, ccpFtn, ccpHdd, ccpAtn, ccpEdn, ccpTxbx, ccpHdrTxbx).
const LW_INDEX = [3, 4, 5, 7, 8, 9, 10] as const;

function resolveLimits(limits?: DocStructureLimits): Required<DocStructureLimits> {
	const out = { ...DEFAULT_DOC_STRUCTURE_LIMITS, ...Object.fromEntries(Object.entries(limits ?? {}).filter(([, v]) => v !== undefined)) } as Required<DocStructureLimits>;
	for (const v of Object.values(out)) {
		if (!Number.isSafeInteger(v) || v < 0) throw new Error('DOC structure limits must be nonnegative safe integers');
	}
	return out;
}

/** Open the document, decode text, and lay out the stories. Throws on encrypted, unsupported or over-budget input. */
export function openStoryContext(input: Uint8Array, limitsIn?: DocStructureLimits): StoryContext {
	const limits = resolveLimits(limitsIn);
	const cfb = unwrapDocBytes(input);
	if (!cfb) throw new Error('Not a readable Word 97-2003 document');
	const word = cfb.wordDocBytes;
	const table = cfb.tableBytes;
	const fib = readDocFib(word);
	const lw = fib.fibRgFcLcbOffset - 2 - 22 * 4;
	const view = new DataView(word.buffer, word.byteOffset, word.byteLength);
	const extents = {} as Record<DocStoryId, StoryExtent>;
	let base = 0;
	ORDER.forEach((id, i) => {
		const length = view.getInt32(lw + LW_INDEX[i]! * 4, true);
		extents[id] = { start: base, length };
		base += length;
	});
	if (base > limits.maxCharacters) throw new DocResourceLimitError('DOC character count exceeds the processing limit');
	const pieces = parseDocClx(table, fib.clx, limits.maxPieces).pieces;
	const text = decodePiecesText(word, pieces, base);
	const warnings: string[] = [];
	const starts = pieces.map((p) => p.cpStart);

	const ctx: StoryContext = {
		word,
		table,
		fib,
		pieces,
		text,
		limits,
		extents,
		warnings,
		pair(index) {
			try {
				return readFcLcbAt(word, fib, index);
			} catch {
				return { fc: 0, lcb: 0 };
			}
		},
		range(story, start, end) {
			const e = extents[story];
			const s = Math.max(0, Math.min(start, e.length));
			const t = Math.max(s, Math.min(end, e.length));
			return { story, start: s, end: t, globalStart: e.start + s, text: text.slice(e.start + s, e.start + t) };
		},
		fcAt(cp) {
			let lo = 0;
			let hi = starts.length;
			while (lo < hi) {
				const mid = (lo + hi) >>> 1;
				if (starts[mid]! <= cp) lo = mid + 1;
				else hi = mid;
			}
			const piece = pieces[lo - 1];
			if (!piece || cp >= piece.cpEnd) return undefined;
			return piece.fc + (cp - piece.cpStart) * (piece.compressed ? 1 : 2);
		},
		plc(index, dataSize) {
			const at = ctx.pair(index);
			if (at.lcb === 0) return undefined;
			if (at.fc > table.length - at.lcb) {
				warnings.push(`PLC ${index} lies outside the table stream`);
				return undefined;
			}
			const stride = 4 + dataSize;
			const n = dataSize === 0 ? at.lcb / 4 - 1 : (at.lcb - 4) / stride;
			if (!Number.isInteger(n) || n < 0) {
				warnings.push(`PLC ${index} has an invalid length`);
				return undefined;
			}
			if (n > limits.maxEntries) throw new DocResourceLimitError('DOC PLC entry count exceeds the processing limit');
			const tv = new DataView(table.buffer, table.byteOffset, table.byteLength);
			const cps: number[] = [];
			for (let i = 0; i <= n; i++) {
				const cp = tv.getInt32(at.fc + i * 4, true);
				if (cp < 0 || (i > 0 && cp < cps[i - 1]!)) {
					warnings.push(`PLC ${index} has non-increasing positions`);
					return undefined;
				}
				cps.push(cp);
			}
			const dataStart = at.fc + (n + 1) * 4;
			return { cps, count: n, dataAt: (i) => dataStart + i * dataSize };
		},
	};
	return ctx;
}

export const STORY_ORDER = ORDER;
