/**
 * Paragraph property lookup by CP for the structure and formatting readers,
 * over the one shared PAPX run index ({@link StoryContext.fkpRuns}). The grpprl
 * of each run is decoded once and cached with the direct paragraph properties
 * ([MS-DOC] 2.9.172, 2.9.173) and the table (sgc 5) sprms.
 *
 * @module doc-papx-index
 */
import type { DocFkpRun } from './doc-fkp-runs.js';
import { applyParaSprms } from './doc-para-props.js';
import type { DocParaProps } from './doc-para-props.js';
import { parseGrpprl, sprmGroup } from './doc-sprm.js';
import type { DocRawSprm } from './doc-sprm.js';
import type { StoryContext } from './doc-story.js';

/** Direct paragraph formatting at one paragraph mark. */
export interface DocParaInfo {
	/** Paragraph style index of the PAPX (0 when the mark is uncovered). */
	readonly istd: number;
	/** The paragraph sprms: PAPX exceptions followed by the piece's PRM exceptions. */
	readonly sprms: readonly DocRawSprm[];
	/** Direct properties only (no style cascade): table flags, nesting level, list and indent overrides. */
	readonly props: DocParaProps;
	/** Table (TAP) sprms carried by this paragraph mark, in order. */
	readonly tap: readonly DocRawSprm[];
}

const NO_INFO: DocParaInfo = Object.freeze({ istd: 0, sprms: [], props: Object.freeze({}), tap: [] });
const cache = new WeakMap<DocFkpRun, DocParaInfo>();

function decode(run: DocFkpRun | undefined, extra: readonly DocRawSprm[], ctx: StoryContext): DocParaInfo {
	const sprms = [...(run ? parseGrpprl(run.grpprl, ctx.budget) : []), ...extra];
	const props: DocParaProps = {};
	applyParaSprms(props, sprms);
	return { istd: run?.istd ?? 0, sprms, props, tap: sprms.filter((s) => sprmGroup(s.opcode) === 5) };
}

/** Paragraph formatting of the paragraph whose mark is at absolute CP `cp`. */
export function paragraphInfoAt(ctx: StoryContext, cp: number): DocParaInfo {
	const piece = ctx.pieceAt(cp);
	if (!piece) return NO_INFO;
	const run = ctx.runAt('papx', piece.fc + (cp - piece.cpStart) * (piece.compressed ? 1 : 2));
	const extra = ctx.pieceSprms(piece);
	if (extra.length > 0 || !run) return decode(run, extra, ctx);
	let hit = cache.get(run);
	if (!hit) cache.set(run, (hit = decode(run, extra, ctx)));
	return hit;
}

/** Table nesting depth of a paragraph: sprmPItap when present, else 1 inside a table and 0 outside. */
export function tableLevel(p: DocParaProps): number {
	return p.itap !== undefined && p.itap >= 0 ? p.itap : p.inTable ? 1 : 0;
}
