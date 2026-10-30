/**
 * Split the main story into paragraphs and runs, mapping CPs to file offsets
 * through the piece table (compressed and Unicode pieces, fast-saved files)
 * and looking formatting up in the sorted PAPX/CHPX ranges.
 *
 * @module doc-formatting-paragraphs
 */
import { applyCharSprms } from './doc-char-props.js';
import type { DocCharProps } from './doc-char-props.js';
import { firstRunEndingAfter } from './doc-fkp-runs.js';
import type { DocFkpRun } from './doc-fkp-runs.js';
import type { DocFormatParagraph, DocFormatRun } from './doc-formatting.js';
import { parseGrpprl } from './doc-sprm.js';
import type { DocRawSprm, SprmBudget } from './doc-sprm.js';
import type { DocStyleResolver } from './doc-style-resolve.js';
import { DocResourceLimitError } from './ole-document-doc-pieces.js';
import type { DocPiece } from './ole-document-doc-pieces.js';

export interface SplitInput {
	text: string;
	pieces: readonly DocPiece[];
	prcBytes: Uint8Array;
	chpx: readonly DocFkpRun[];
	papx: readonly DocFkpRun[];
	resolver: DocStyleResolver;
	budget: SprmBudget;
	maxParagraphs: number;
	maxRuns: number;
}

/** Prc records (clxt 1, cb, grpprl) in order; Prm1 selects one by index. */
function parsePrcs(bytes: Uint8Array, budget: SprmBudget): DocRawSprm[][] {
	const out: DocRawSprm[][] = [];
	for (let at = 0; at + 3 <= bytes.length;) {
		const cb = bytes[at + 1]! | (bytes[at + 2]! << 8);
		out.push(parseGrpprl(bytes.subarray(at + 3, at + 3 + cb), budget));
		at += 3 + cb;
	}
	return out;
}

function deepFreeze<T extends DocCharProps>(p: T): T {
	if (p.fontIds) Object.freeze(p.fontIds);
	if (p.fonts) Object.freeze(p.fonts);
	return Object.freeze(p);
}

const EMPTY: Uint8Array = new Uint8Array(0);

export function splitParagraphs(input: SplitInput): DocFormatParagraph[] {
	const { text, pieces, chpx, papx, resolver, budget } = input;
	const prcs = parsePrcs(input.prcBytes, budget);
	const prcOf = (piece: DocPiece): DocRawSprm[] => ((piece.prm & 1) === 1 ? prcs[piece.prm >>> 1] ?? [] : []);
	const grpprlCache = new Map<DocFkpRun | null, DocRawSprm[]>();
	const grpprlOf = (run: DocFkpRun | undefined): DocRawSprm[] => {
		const key = run ?? null;
		let hit = grpprlCache.get(key);
		if (!hit) grpprlCache.set(key, (hit = parseGrpprl(run ? run.grpprl : EMPTY, budget)));
		return hit;
	};
	const propsCache = new Map<DocFkpRun | null, Map<string, Readonly<DocCharProps>>>();
	let runCount = 0;

	/** Resolved character properties for a CHPX range under a paragraph style and piece. */
	const charProps = (run: DocFkpRun | undefined, paraStyle: number, piece: DocPiece): Readonly<DocCharProps> => {
		const key = run ?? null;
		let inner = propsCache.get(key);
		if (!inner) propsCache.set(key, (inner = new Map()));
		const id = `${paraStyle}:${prcs.length && (piece.prm & 1) ? piece.prm : 0}`;
		const hit = inner.get(id);
		if (hit) return hit;
		const sprms = [...grpprlOf(run), ...prcOf(piece)];
		let cstyle: number | undefined;
		for (const s of sprms) if (s.opcode === 0x4a30 && s.operand.length >= 2) cstyle = s.operand[0]! | (s.operand[1]! << 8);
		const props = resolver.baseCharProps(paraStyle, cstyle);
		applyCharSprms(props, sprms);
		if (cstyle !== undefined) props.styleId = cstyle;
		resolver.withFontNames(props);
		const frozen = deepFreeze(props);
		inner.set(id, frozen);
		return frozen;
	};

	const paragraphs: DocFormatParagraph[] = [];
	let pieceIdx = 0;
	const pieceFor = (cp: number): DocPiece | undefined => {
		while (pieceIdx < pieces.length && pieces[pieceIdx]!.cpEnd <= cp) pieceIdx++;
		const p = pieces[pieceIdx];
		return p && p.cpStart <= cp ? p : undefined;
	};
	const fcOf = (piece: DocPiece, cp: number): number => piece.fc + (cp - piece.cpStart) * (piece.compressed ? 1 : 2);
	const runAt = (list: readonly DocFkpRun[], fc: number): DocFkpRun | undefined => {
		const r = list[firstRunEndingAfter(list, fc)];
		return r && r.fcStart <= fc ? r : undefined;
	};

	let start = 0;
	for (let cp = 0; cp < text.length; cp++) {
		const code = text.charCodeAt(cp);
		if (code !== 13 && code !== 7) continue;
		const piece = pieceFor(cp);
		if (!piece) break;
		const prun = runAt(papx, fcOf(piece, cp));
		const sprms = [...grpprlOf(prun), ...prcOf(piece)];
		const props = resolver.paragraphProps(prun?.istd ?? 0, sprms);
		// A bell is only a paragraph mark inside tables; elsewhere it is ordinary text.
		if (code === 7 && !props.inTable && !props.tableRowEnd) continue;
		if (paragraphs.length >= input.maxParagraphs) throw new DocResourceLimitError('DOC paragraph count exceeds the processing limit');
		paragraphs.push(build(start, cp, text[cp]!, props, charProps(runAt(chpx, fcOf(piece, cp)), props.styleId ?? 0, piece)));
		start = cp + 1;
	}
	if (start < text.length) {
		const last = pieceFor(text.length - 1);
		const prun = last ? runAt(papx, fcOf(last, text.length - 1)) : undefined;
		const props = resolver.paragraphProps(prun?.istd ?? 0, last ? [...grpprlOf(prun), ...prcOf(last)] : []);
		const piece = last ?? pieces[0]!;
		paragraphs.push(build(start, text.length, '', props, charProps(undefined, props.styleId ?? 0, piece)));
	}
	return paragraphs;

	function build(s: number, end: number, terminator: string, props: DocFormatParagraph['props'], mark: Readonly<DocCharProps>): DocFormatParagraph {
		const style = props.styleId ?? 0;
		const runs = runsFor(s, end, style);
		return { cpStart: s, cpEnd: terminator ? end + 1 : end, text: text.slice(s, end), style, props, markProps: mark, terminator, runs };
	}

	function runsFor(s: number, end: number, paraStyle: number): DocFormatRun[] {
		const runs: DocFormatRun[] = [];
		let cp = s;
		while (cp < end) {
			const piece = pieceFor(cp);
			if (!piece) break;
			const unit = piece.compressed ? 1 : 2;
			const stop = Math.min(end, piece.cpEnd);
			let idx = firstRunEndingAfter(chpx, fcOf(piece, cp));
			while (cp < stop) {
				const fc = fcOf(piece, cp);
				const run = chpx[idx];
				let next: number;
				let used: DocFkpRun | undefined;
				if (!run || run.fcStart > fc) {
					next = run ? Math.min(stop, cp + Math.max(1, Math.ceil((run.fcStart - fc) / unit))) : stop;
				} else {
					used = run;
					next = Math.min(stop, cp + Math.max(1, Math.ceil((run.fcEnd - fc) / unit)));
					if (fcOf(piece, next) >= run.fcEnd) idx++;
				}
				const props = charProps(used, paraStyle, piece);
				const prev = runs.at(-1);
				if (prev && prev.props === props && prev.cpEnd === cp) prev.cpEnd = next;
				else {
					if (++runCount > input.maxRuns) throw new DocResourceLimitError('DOC run count exceeds the processing limit');
					runs.push({ cpStart: cp, cpEnd: next, text: '', style: props.styleId, props });
				}
				cp = next;
			}
		}
		for (const r of runs) r.text = text.slice(r.cpStart, r.cpEnd);
		return runs;
	}
}
