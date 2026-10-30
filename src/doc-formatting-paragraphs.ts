/**
 * Split one story into paragraphs and runs, mapping CPs to file offsets
 * through the shared piece table (compressed and Unicode pieces, fast-saved
 * files) and the shared PAPX/CHPX run indexes of the {@link StoryContext}.
 *
 * @module doc-formatting-paragraphs
 */
import { applyCharSprms } from './doc-char-props.js';
import type { DocCharProps } from './doc-char-props.js';
import { firstRunEndingAfter } from './doc-fkp-runs.js';
import type { DocFkpRun } from './doc-fkp-runs.js';
import type { DocFormatParagraph, DocFormatRun } from './doc-formatting.js';
import { paragraphInfoAt } from './doc-papx-index.js';
import { parseGrpprl } from './doc-sprm.js';
import type { DocRawSprm } from './doc-sprm.js';
import type { StoryContext } from './doc-story.js';
import type { DocStyleResolver } from './doc-style-resolve.js';
import { DocResourceLimitError } from './ole-document-doc-pieces.js';
import type { DocPiece } from './ole-document-doc-pieces.js';

/** Counters shared across every story of one read, so the limits bound the whole document. */
export interface SplitCounters {
	paragraphs: number;
	runs: number;
}

export interface SplitInput {
	ctx: StoryContext;
	resolver: DocStyleResolver;
	/** Absolute CP range of the story to split. */
	cpStart: number;
	cpEnd: number;
	/** Absolute CPs just past a section mark; a form feed there ends a paragraph (main story only). */
	sectionEnds: ReadonlySet<number>;
	counters: SplitCounters;
}

function deepFreeze<T extends DocCharProps>(p: T): T {
	if (p.fontIds) Object.freeze(p.fontIds);
	if (p.fonts) Object.freeze(p.fonts);
	return Object.freeze(p);
}

const EMPTY: Uint8Array = new Uint8Array(0);

export function splitParagraphs(input: SplitInput): DocFormatParagraph[] {
	const { ctx, resolver, counters } = input;
	const text = ctx.text;
	const chpx = ctx.fkpRuns('chpx');
	const grpprlCache = new Map<DocFkpRun | null, readonly DocRawSprm[]>();
	const grpprlOf = (run: DocFkpRun | undefined): readonly DocRawSprm[] => {
		const key = run ?? null;
		let hit = grpprlCache.get(key);
		if (!hit) grpprlCache.set(key, (hit = parseGrpprl(run ? run.grpprl : EMPTY, ctx.budget)));
		return hit;
	};
	const propsCache = new Map<DocFkpRun | null, Map<string, Readonly<DocCharProps>>>();

	/** Resolved character properties for a CHPX range under a paragraph style and piece. */
	const charProps = (run: DocFkpRun | undefined, paraStyle: number, piece: DocPiece): Readonly<DocCharProps> => {
		const key = run ?? null;
		let inner = propsCache.get(key);
		if (!inner) propsCache.set(key, (inner = new Map()));
		const id = `${paraStyle}:${piece.prm & 1 ? piece.prm : 0}`;
		const hit = inner.get(id);
		if (hit) return hit;
		const sprms = [...grpprlOf(run), ...ctx.pieceSprms(piece)];
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

	const fcOf = (piece: DocPiece, cp: number): number => piece.fc + (cp - piece.cpStart) * (piece.compressed ? 1 : 2);
	const chpxAt = (fc: number): DocFkpRun | undefined => ctx.runAt('chpx', fc);
	const isMain = ctx.storyOf(input.cpStart) === 'main';
	const end = Math.min(input.cpEnd, text.length);
	const paragraphs: DocFormatParagraph[] = [];

	const paraProps = (cp: number) => {
		const info = paragraphInfoAt(ctx, cp);
		return resolver.paragraphProps(info.istd, info.sprms);
	};
	const count = () => {
		if (++counters.paragraphs > ctx.limits.maxParagraphs) throw new DocResourceLimitError('DOC paragraph count exceeds the processing limit');
	};

	let start = input.cpStart;
	for (let cp = start; cp < end; cp++) {
		const code = text.charCodeAt(cp);
		const formFeed = code === 12 && isMain && input.sectionEnds.has(cp + 1);
		if (code !== 13 && code !== 7 && !formFeed) continue;
		const piece = ctx.pieceAt(cp);
		if (!piece) break;
		const props = paraProps(cp);
		// A bell is only a paragraph mark inside tables; elsewhere it is ordinary text.
		if (code === 7 && !props.inTable && !props.tableRowEnd) continue;
		count();
		paragraphs.push(build(start, cp, text[cp]!, props, charProps(chpxAt(fcOf(piece, cp)), props.styleId ?? 0, piece)));
		start = cp + 1;
	}
	if (start < end) {
		const last = ctx.pieceAt(end - 1);
		const props = last ? paraProps(end - 1) : resolver.paragraphProps(0, []);
		const piece = last ?? ctx.pieces[0];
		if (piece) {
			count();
			paragraphs.push(build(start, end, '', props, charProps(undefined, props.styleId ?? 0, piece)));
		}
	}
	return paragraphs;

	function build(s: number, e: number, terminator: string, props: DocFormatParagraph['props'], mark: Readonly<DocCharProps>): DocFormatParagraph {
		const style = props.styleId ?? 0;
		const runs = runsFor(s, e, style);
		return { cpStart: s, cpEnd: terminator ? e + 1 : e, text: text.slice(s, e), style, props, markProps: mark, terminator, runs };
	}

	function runsFor(s: number, e: number, paraStyle: number): DocFormatRun[] {
		const runs: DocFormatRun[] = [];
		let cp = s;
		while (cp < e) {
			const piece = ctx.pieceAt(cp);
			if (!piece) break;
			const unit = piece.compressed ? 1 : 2;
			const stop = Math.min(e, piece.cpEnd);
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
					if (++counters.runs > ctx.limits.maxRuns) throw new DocResourceLimitError('DOC run count exceeds the processing limit');
					runs.push({ cpStart: cp, cpEnd: next, text: '', style: props.styleId, props });
				}
				cp = next;
			}
		}
		for (const r of runs) r.text = text.slice(r.cpStart, r.cpEnd);
		return runs;
	}
}
