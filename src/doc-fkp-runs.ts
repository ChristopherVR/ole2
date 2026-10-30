/**
 * Flatten PlcBteChpx/PlcBtePapx and the formatted disk pages (FKPs) they
 * reference into sorted FC ranges with their raw grpprl ([MS-DOC] 2.9.40,
 * 2.9.42, 2.9.69, 2.9.162).
 *
 * @module doc-fkp-runs
 */
import { DocResourceLimitError } from './ole-document-doc-pieces.js';

const PAGE = 512;

/** One formatted range of the `WordDocument` stream. */
export interface DocFkpRun {
	/** First byte offset covered (inclusive). */
	fcStart: number;
	/** Byte offset just past the range. */
	fcEnd: number;
	/** Raw sprms (a view into the stream); empty when the range carries no exceptions. */
	grpprl: Uint8Array;
	/** Paragraph style index for PAPX ranges. */
	istd?: number;
}

/** Page and run budgets. */
export interface FkpLimits {
	maxPages: number;
	maxRuns: number;
}

function view(b: Uint8Array): DataView {
	return new DataView(b.buffer, b.byteOffset, b.byteLength);
}

/** Parse a PlcBte at `(fc, lcb)` into its page numbers; bounded before any allocation. */
function readPageNumbers(table: Uint8Array, at: { fc: number; lcb: number }, limits: FkpLimits): number[] {
	if (at.lcb === 0) return [];
	if (at.lcb < 12 || (at.lcb - 4) % 8 !== 0 || at.fc > table.length - at.lcb)
		throw new Error('Invalid formatting BTE table bounds');
	const n = (at.lcb - 4) / 8;
	if (n > limits.maxPages) throw new DocResourceLimitError('DOC formatting page count exceeds the processing limit');
	const v = view(table);
	const out: number[] = [];
	for (let i = 0; i < n; i++) out.push(v.getUint32(at.fc + (n + 1) * 4 + i * 4, true) & 0x3fffff);
	return out;
}

/** Read every CHPX or PAPX run, sorted by `fcStart`; overlapping or inverted ranges are dropped. */
export function readFkpRuns(
	word: Uint8Array,
	table: Uint8Array,
	at: { fc: number; lcb: number },
	kind: 'chpx' | 'papx',
	limits: FkpLimits,
): DocFkpRun[] {
	const pages = readPageNumbers(table, at, limits);
	const v = view(word);
	const out: DocFkpRun[] = [];
	for (const pn of pages) {
		const base = pn * PAGE;
		if (base > word.length - PAGE) throw new Error('Formatting page outside WordDocument');
		const crun = word[base + PAGE - 1]!;
		const rowSize = kind === 'chpx' ? 1 : 13;
		const header = (crun + 1) * 4 + crun * rowSize;
		if (crun === 0 || header > PAGE - 1) continue;
		if (out.length + crun > limits.maxRuns) throw new DocResourceLimitError('DOC formatting run count exceeds the processing limit');
		for (let i = 0; i < crun; i++) {
			const fcStart = v.getInt32(base + i * 4, true);
			const fcEnd = v.getInt32(base + (i + 1) * 4, true);
			if (fcStart < 0 || fcEnd <= fcStart) continue;
			const word2 = word[base + (crun + 1) * 4 + i * rowSize]! * 2;
			const run = word2 === 0 ? emptyRun(fcStart, fcEnd, kind) : decodeRun(word, base, word2, header, fcStart, fcEnd, kind);
			if (run) out.push(run);
		}
	}
	out.sort((a, b) => a.fcStart - b.fcStart);
	const clean: DocFkpRun[] = [];
	for (const run of out) {
		const last = clean.at(-1);
		if (last && run.fcStart < last.fcEnd) continue;
		clean.push(run);
	}
	return clean;
}

function emptyRun(fcStart: number, fcEnd: number, kind: 'chpx' | 'papx'): DocFkpRun {
	const run: DocFkpRun = { fcStart, fcEnd, grpprl: new Uint8Array(0) };
	if (kind === 'papx') run.istd = 0;
	return run;
}

function decodeRun(
	word: Uint8Array, base: number, offset: number, header: number, fcStart: number, fcEnd: number, kind: 'chpx' | 'papx',
): DocFkpRun | undefined {
	if (offset < header || offset >= PAGE - 1) return undefined;
	const cb = word[base + offset]!;
	if (kind === 'chpx') {
		if (offset + 1 + cb > PAGE - 1) return undefined;
		return { fcStart, fcEnd, grpprl: word.subarray(base + offset + 1, base + offset + 1 + cb) };
	}
	// PapxInFkp: cb counts words; cb == 0 defers the size to the next byte. The grpprl starts with the istd.
	const bytes = cb === 0 ? word[base + offset + 1]! * 2 : cb * 2 - 1;
	const start = base + offset + (cb === 0 ? 2 : 1);
	if (offset + (cb === 0 ? 2 : 1) + bytes > PAGE - 1 || bytes < 2) return undefined;
	const istd = word[start]! | (word[start + 1]! << 8);
	return { fcStart, fcEnd, grpprl: word.subarray(start + 2, start + bytes), istd };
}

/** Index of the first run whose `fcEnd` is greater than `fc` (binary search). */
export function firstRunEndingAfter(runs: readonly DocFkpRun[], fc: number): number {
	let lo = 0, hi = runs.length;
	while (lo < hi) {
		const mid = (lo + hi) >>> 1;
		if (runs[mid]!.fcEnd <= fc) lo = mid + 1;
		else hi = mid;
	}
	return lo;
}
