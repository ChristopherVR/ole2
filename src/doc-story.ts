/**
 * The one shared context every `.doc` reader opens: streams, FIB, piece table,
 * decoded text, the story layout of the document-wide character position (CP)
 * space ([MS-DOC] 2.4.1, FibRgLw97), bounded PLC access and the shared FKP and
 * sprm lookups.
 *
 * CP convention: a CP is always an absolute document CP unless a field is
 * documented as story-relative. A {@link DocStoryRange} carries both: the story
 * id, `cpStart`/`cpEnd` (absolute) and `start`/`end` (relative to the story).
 *
 * @module doc-story
 */
import { firstRunEndingAfter, readFkpRuns } from './doc-fkp-runs.js';
import type { DocFkpRun } from './doc-fkp-runs.js';
import { resolveDocLimits } from './doc-limits.js';
import type { DocModelLimits, ResolvedDocLimits } from './doc-limits.js';
import { parseGrpprl } from './doc-sprm.js';
import type { DocRawSprm, SprmBudget } from './doc-sprm.js';
import { unwrapDocBytes } from './ole-document-doc-cfb.js';
import { readDocFib, readFcLcbAt } from './ole-document-doc-fib.js';
import type { FcLcb, ParsedDocFib } from './ole-document-doc-fib.js';
import { DocResourceLimitError, decodePiecesText, parseDocClx } from './ole-document-doc-pieces.js';
import type { DocPiece } from './ole-document-doc-pieces.js';
import { readCompoundFileStream } from './ole2-stream-edit.js';

/** The seven stories of a Word document, in CP order. */
export type DocStoryId = 'main' | 'footnotes' | 'headers' | 'comments' | 'endnotes' | 'textboxes' | 'headerTextboxes';

/** A character range in one story. */
export interface DocStoryRange {
	readonly story: DocStoryId;
	/** Absolute document CP of the first character. */
	readonly cpStart: number;
	/** Absolute document CP just past the last character. */
	readonly cpEnd: number;
	/** Story-relative start (`cpStart` minus the story's base). */
	readonly start: number;
	/** Story-relative end, exclusive. */
	readonly end: number;
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
	readonly input: Uint8Array;
	readonly word: Uint8Array;
	readonly table: Uint8Array;
	/** The `Data` stream (inline picture and field data), when present. */
	readonly data: Uint8Array | undefined;
	readonly fib: ParsedDocFib;
	readonly pieces: readonly DocPiece[];
	readonly prcBytes: Uint8Array;
	/** Text of every story, concatenated, indexed by absolute CP. */
	readonly text: string;
	readonly limits: ResolvedDocLimits;
	readonly extents: Readonly<Record<DocStoryId, StoryExtent>>;
	/** The non-empty stories, in CP order. */
	readonly stories: readonly DocStoryRange[];
	readonly warnings: string[];
	/** Sprm budget shared by everything read through this context. */
	readonly budget: SprmBudget;
	/** Raw FibRgFcLcb97/2000/2002 pair, or empty when the index is beyond this FIB's array. */
	pair(index: number): FcLcb;
	range(story: DocStoryId, start: number, end: number): DocStoryRange;
	/** The story an absolute CP lies in (`main` when outside every story). */
	storyOf(cp: number): DocStoryId;
	/** Byte offset in WordDocument of absolute CP `cp`, or undefined if no piece covers it. */
	fcAt(cp: number): number | undefined;
	plc(index: number, dataSize: number): Plc | undefined;
	/** All CHPX or PAPX runs in file order, read once. */
	fkpRuns(kind: 'chpx' | 'papx'): readonly DocFkpRun[];
	/** The FKP run of `kind` covering byte `fc`. */
	runAt(kind: 'chpx' | 'papx', fc: number): DocFkpRun | undefined;
	/** The piece covering absolute CP `cp`. */
	pieceAt(cp: number): DocPiece | undefined;
	/** The grpprl of the piece's Prm1 exception, parsed once. */
	pieceSprms(piece: DocPiece): readonly DocRawSprm[];
	/** Direct character sprms at a CP: CHPX exceptions followed by piece PRM exceptions. */
	sprmsAt(cp: number): readonly DocRawSprm[];
}

export const STORY_ORDER: readonly DocStoryId[] = ['main', 'footnotes', 'headers', 'comments', 'endnotes', 'textboxes', 'headerTextboxes'];
// FibRgLw97 element index of each story length (ccpText, ccpFtn, ccpHdd, ccpAtn, ccpEdn, ccpTxbx, ccpHdrTxbx).
const LW_INDEX = [3, 4, 5, 7, 8, 9, 10] as const;
const NONE: FcLcb = { fc: 0, lcb: 0 };

/** Grpprl bytes of each Prc in the CLX (Prm1 selects one by index). Damaged entries end the list. */
function prcGrpprls(prcBytes: Uint8Array): Uint8Array[] {
	const out: Uint8Array[] = [];
	for (let at = 0; at + 3 <= prcBytes.length && prcBytes[at] === 1; ) {
		const cb = prcBytes[at + 1]! | (prcBytes[at + 2]! << 8);
		if (at + 3 + cb > prcBytes.length) break;
		out.push(prcBytes.subarray(at + 3, at + 3 + cb));
		at += 3 + cb;
	}
	return out;
}

/** Open the document, decode text, and lay out the stories. Throws on encrypted, unsupported or over-budget input. */
export function openStoryContext(input: Uint8Array, limitsIn?: DocModelLimits): StoryContext {
	const limits = resolveDocLimits(limitsIn);
	const cfb = unwrapDocBytes(input);
	if (!cfb) throw new Error('Not a readable Word 97-2003 document');
	const word = cfb.wordDocBytes;
	const table = cfb.tableBytes;
	const fib = readDocFib(word);
	const lw = fib.fibRgFcLcbOffset - 2 - 22 * 4;
	const view = new DataView(word.buffer, word.byteOffset, word.byteLength);
	const extents = {} as Record<DocStoryId, StoryExtent>;
	let base = 0;
	STORY_ORDER.forEach((id, i) => {
		const length = view.getInt32(lw + LW_INDEX[i]! * 4, true);
		if (length < 0) throw new Error('Negative story length');
		extents[id] = { start: base, length };
		base += length;
	});
	if (base > limits.maxCharacters) throw new DocResourceLimitError('DOC character count exceeds the processing limit');
	const clx = parseDocClx(table, fib.clx, limits.maxPieces);
	const pieces = clx.pieces;
	const total = Math.max(base, pieces.at(-1)?.cpEnd ?? 0);
	if (total > limits.maxCharacters) throw new DocResourceLimitError('DOC text exceeds the processing limit');
	const text = decodePiecesText(word, pieces, total);
	const warnings: string[] = [];
	const budget: SprmBudget = { remaining: limits.maxSprms };
	const starts = pieces.map((p) => p.cpStart);
	const prcs = prcGrpprls(clx.prcBytes);
	const prcCache = new Map<number, readonly DocRawSprm[]>();
	const runCache = new Map<string, readonly DocFkpRun[]>();
	const sprmCache = new Map<number, readonly DocRawSprm[]>();

	const range = (story: DocStoryId, start: number, end: number): DocStoryRange => {
		const e = extents[story];
		const s = Math.max(0, Math.min(start, e.length));
		const t = Math.max(s, Math.min(end, e.length));
		return Object.freeze({ story, cpStart: e.start + s, cpEnd: e.start + t, start: s, end: t, text: text.slice(e.start + s, e.start + t) });
	};
	const stories = STORY_ORDER.filter((id) => extents[id].length > 0).map((id) => range(id, 0, extents[id].length));

	const ctx: StoryContext = {
		input,
		word,
		table,
		data: readCompoundFileStream(input, ['Data']),
		fib,
		pieces,
		prcBytes: clx.prcBytes,
		text,
		limits,
		extents,
		stories,
		warnings,
		budget,
		pair(index) {
			try {
				return readFcLcbAt(word, fib, index);
			} catch {
				return NONE;
			}
		},
		range,
		storyOf(cp) {
			for (const s of stories) if (cp >= s.cpStart && cp < s.cpEnd) return s.story;
			return 'main';
		},
		pieceAt(cp) {
			let lo = 0;
			let hi = starts.length;
			while (lo < hi) {
				const mid = (lo + hi) >>> 1;
				if (starts[mid]! <= cp) lo = mid + 1;
				else hi = mid;
			}
			const piece = pieces[lo - 1];
			return piece && cp < piece.cpEnd ? piece : undefined;
		},
		fcAt(cp) {
			const piece = ctx.pieceAt(cp);
			return piece ? piece.fc + (cp - piece.cpStart) * (piece.compressed ? 1 : 2) : undefined;
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
		fkpRuns(kind) {
			let runs = runCache.get(kind);
			if (!runs) {
				try {
					const at = kind === 'chpx' ? fib.plcfbteChpx : fib.plcfbtePapx;
					runs = readFkpRuns(word, table, at, kind, { maxPages: limits.maxFormattingPages, maxRuns: limits.maxRuns });
				} catch (e) {
					if (e instanceof DocResourceLimitError) throw e;
					warnings.push(`${kind === 'chpx' ? 'Character' : 'Paragraph'} formatting table unreadable: ${(e as Error).message}`);
					runs = [];
				}
				runCache.set(kind, runs);
			}
			return runs;
		},
		runAt(kind, fc) {
			const list = ctx.fkpRuns(kind);
			const r = list[firstRunEndingAfter(list, fc)];
			return r && r.fcStart <= fc ? r : undefined;
		},
		pieceSprms(piece) {
			if ((piece.prm & 1) !== 1) return [];
			const index = piece.prm >>> 1;
			let hit = prcCache.get(index);
			if (!hit) prcCache.set(index, (hit = prcs[index] ? parseGrpprl(prcs[index]!, budget) : []));
			return hit;
		},
		sprmsAt(cp) {
			const cached = sprmCache.get(cp);
			if (cached) return cached;
			const piece = ctx.pieceAt(cp);
			let out: readonly DocRawSprm[] = [];
			if (piece) {
				const fc = piece.fc + (cp - piece.cpStart) * (piece.compressed ? 1 : 2);
				const run = ctx.runAt('chpx', fc);
				out = [...(run ? parseGrpprl(run.grpprl, budget) : []), ...ctx.pieceSprms(piece)];
			}
			if (sprmCache.size < 65_536) sprmCache.set(cp, out);
			return out;
		},
	};
	return ctx;
}
