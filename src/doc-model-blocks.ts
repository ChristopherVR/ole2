/**
 * Build the block tree of a story range: paragraphs (with list references and
 * inline items) and tables whose cells hold blocks, nested to any depth up to
 * the `maxNesting` budget.
 *
 * @module doc-model-blocks
 */
import type { DocFormatParagraph } from './doc-formatting.js';
import { buildInlines } from './doc-model-inline.js';
import type { FieldState, InlineEnv } from './doc-model-inline.js';
import type { DocBlock, DocParagraphBlock, DocParagraphList, DocTableBlock, DocTableCellBlock } from './doc-model-types.js';
import type { DocLists } from './doc-lists.js';
import { gridBoundaries } from './doc-tables.js';
import type { DocTable } from './doc-tables.js';
import { DocResourceLimitError } from './ole-document-doc-pieces.js';

export interface BlockEnv {
	readonly inline: InlineEnv;
	readonly lists: DocLists;
	/** Absolute paragraph-end CP to the index of the section that paragraph ends. */
	readonly sectionEnds: ReadonlyMap<number, number>;
	/** Absolute CP of each story's first character, by story, to turn story-relative table CPs absolute. */
	readonly storyBase: (story: DocTable['story']) => number;
}

/** Index of the first paragraph whose `cpStart` is at least `cp`. */
function lowerBound(paras: readonly DocFormatParagraph[], cp: number): number {
	let lo = 0;
	let hi = paras.length;
	while (lo < hi) {
		const mid = (lo + hi) >>> 1;
		if (paras[mid]!.cpStart < cp) lo = mid + 1;
		else hi = mid;
	}
	return lo;
}

function listOf(env: BlockEnv, p: DocFormatParagraph): DocParagraphList | undefined {
	const ilfo = p.props.ilfo ?? 0;
	if (ilfo <= 0) return undefined;
	const level = Math.min(Math.max(p.props.ilvl ?? 0, 0), 8);
	const override = env.lists.overrides.find((o) => o.ilfo === ilfo);
	const definition = override?.definitionIndex === undefined ? undefined : env.lists.definitions[override.definitionIndex];
	const levelDefinition = override?.levels.find((l) => l.level === level)?.definition ?? definition?.levels[level];
	return { ilfo, level, override, definition, levelDefinition };
}

function paragraphBlock(env: BlockEnv, p: DocFormatParagraph, state: FieldState): DocParagraphBlock {
	return {
		type: 'paragraph',
		cpStart: p.cpStart,
		cpEnd: p.cpEnd,
		style: p.style,
		props: p.props,
		markProps: p.markProps,
		list: listOf(env, p),
		inlines: buildInlines(env.inline, p, state),
		endsSection: p.terminator ? env.sectionEnds.get(p.cpEnd) : undefined,
	};
}

/**
 * Blocks for `paras[from..to)`. `tables` are the tables that can occur in that
 * range (sorted by position); a paragraph inside a table is consumed by it.
 */
export function buildBlocks(
	env: BlockEnv,
	paras: readonly DocFormatParagraph[],
	from: number,
	to: number,
	tables: readonly DocTable[],
	state: FieldState,
	depth: number,
): DocBlock[] {
	if (depth > env.inline.ctx.limits.maxNesting) throw new DocResourceLimitError('DOC table nesting exceeds the processing limit');
	const out: DocBlock[] = [];
	let ti = 0;
	for (let i = from; i < to; ) {
		const p = paras[i]!;
		let table: DocTable | undefined;
		while (ti < tables.length) {
			const t = tables[ti]!;
			const base = env.storyBase(t.story);
			if (base + t.end <= p.cpStart) ti++;
			else {
				if (base + t.start <= p.cpStart) table = t;
				break;
			}
		}
		if (!table) {
			out.push(paragraphBlock(env, p, state));
			i++;
			continue;
		}
		const base = env.storyBase(table.story);
		out.push(tableBlock(env, paras, table, base, state, depth));
		ti++;
		while (i < to && paras[i]!.cpStart < base + table.end) i++;
	}
	return out;
}

function tableBlock(
	env: BlockEnv,
	paras: readonly DocFormatParagraph[],
	table: DocTable,
	base: number,
	state: FieldState,
	depth: number,
): DocTableBlock {
	return {
		type: 'table',
		level: table.level,
		cpStart: base + table.start,
		cpEnd: base + table.end,
		gridTwips: gridBoundaries(table.rows),
		rows: table.rows.map((row) => ({
			cpStart: base + row.start,
			props: row.props,
			cells: row.cells.map((cell): DocTableCellBlock => {
				const from = lowerBound(paras, base + cell.start);
				const to = lowerBound(paras, base + cell.markCp + 1);
				return {
					gridColumn: cell.gridColumn,
					columnSpan: cell.columnSpan,
					rowSpan: cell.rowSpan,
					vMerge: cell.props?.vertMerge,
					props: cell.props,
					blocks: buildBlocks(env, paras, from, to, cell.nestedTables, state, depth + 1),
				};
			}),
		})),
	};
}

/** Blocks for the paragraphs of `paras` inside the absolute CP range, with that range's tables. */
export function buildRangeBlocks(
	env: BlockEnv,
	paras: readonly DocFormatParagraph[],
	cpStart: number,
	cpEnd: number,
	tables: readonly DocTable[],
	state: FieldState,
): DocBlock[] {
	const from = lowerBound(paras, cpStart);
	const to = lowerBound(paras, cpEnd);
	const inRange = tables.filter((t) => {
		const b = env.storyBase(t.story);
		return b + t.start >= cpStart && b + t.start < cpEnd;
	});
	return buildBlocks(env, paras, from, to, inRange, state, 1);
}
