/**
 * Table detection for a legacy `.doc`: rows, cells and nesting recovered from
 * paragraph marks and their table sprms ([MS-DOC] 2.4.3 table structure).
 * A cell ends with a cell mark (0x07, or a CR flagged sprmPFInnerTableCell in
 * a nested table); a row ends with a table-terminating paragraph (sprmPFTtp /
 * sprmPFInnerTtp) that carries the row's TAP. Positions are story-relative.
 *
 * @module doc-tables
 */
import { DocResourceLimitError } from './ole-document-doc-pieces.js';
import { buildPapxIndex, type DocParaProps } from './doc-papx-index.js';
import { readRowProps, type DocCellProps, type DocRowProps } from './doc-table-props.js';
import { STORY_ORDER, type DocStoryId, type StoryContext } from './doc-story.js';
import type { SprmBudget } from './doc-sprm-reader.js';

export interface DocTableCell {
	/** Story-relative CP of the cell's first character. */
	readonly start: number;
	/** Story-relative CP just past the cell mark. */
	readonly end: number;
	/** Story-relative CP of the cell mark itself (content is `[start, markCp)`). */
	readonly markCp: number;
	/** Content text, without the cell mark (nested tables included as raw characters). */
	readonly text: string;
	/** First grid column (0-based) the cell occupies. */
	readonly gridColumn: number;
	readonly columnSpan: number;
	/** Rows covered when this cell restarts a vertical merge, else 1. */
	readonly rowSpan: number;
	readonly props: DocCellProps | undefined;
	readonly nestedTables: readonly DocTable[];
}

export interface DocTableRow {
	readonly start: number;
	readonly end: number;
	/** Story-relative CP of the row-end mark. */
	readonly markCp: number;
	readonly cells: readonly DocTableCell[];
	/** Undefined when the row has no sprmTDefTable (malformed or truncated table). */
	readonly props: DocRowProps | undefined;
}

export interface DocTable {
	readonly story: DocStoryId;
	/** 1 for an outermost table, 2 and up for tables nested in a cell. */
	readonly level: number;
	readonly start: number;
	readonly end: number;
	readonly rows: readonly DocTableRow[];
}

interface RowBuilder {
	cells: DocTableCell[];
	start: number;
}
interface TableBuilder {
	level: number;
	start: number;
	end: number;
	rows: DocTableRow[];
	row: RowBuilder | undefined;
	cellStart: number | undefined;
	pendingNested: DocTable[];
}

interface Mark {
	kind: 'text' | 'cell' | 'row';
	level: number;
	props: DocParaProps | undefined;
}

function classify(ch: string, p: DocParaProps | undefined): Mark {
	if (!p || !p.inTable) return { kind: 'text', level: 0, props: p };
	const level = Math.max(1, p.itap);
	if (p.rowEnd || p.innerRowEnd) return { kind: 'row', level, props: p };
	if (p.innerCell || ch === '\x07') return { kind: 'cell', level, props: p };
	return { kind: 'text', level, props: p };
}

/** Find every table in every story. `sectionEnds` are main-story CPs just past a section mark. */
export function readDocTables(ctx: StoryContext, sectionEnds: ReadonlySet<number>, budget: SprmBudget): DocTable[] {
	const hasContent = STORY_ORDER.some((s) => ctx.extents[s].length > 0);
	if (!hasContent) return [];
	const papx = buildPapxIndex(ctx.word, ctx.table, ctx.fib.plcfbtePapx, ctx.limits.maxPieces, budget);
	const tables: DocTable[] = [];
	const counters = { paragraphs: 0, cells: 0 };
	for (const story of STORY_ORDER) {
		const ext = ctx.extents[story];
		const end = Math.min(ext.start + ext.length, ctx.text.length);
		if (end <= ext.start) continue;
		scanStory(ctx, papx, story, ext.start, end, sectionEnds, counters, tables);
	}
	return tables;
}

function scanStory(
	ctx: StoryContext,
	papx: ReturnType<typeof buildPapxIndex>,
	story: DocStoryId,
	base: number,
	end: number,
	sectionEnds: ReadonlySet<number>,
	counters: { paragraphs: number; cells: number },
	out: DocTable[],
): void {
	const stack: TableBuilder[] = [];
	const rel = (cp: number) => cp - base;
	const text = ctx.text;
	const closeTo = (depth: number) => {
		while (stack.length > depth) {
			const t = finish(story, stack.pop()!);
			const parent = stack[stack.length - 1];
			if (parent) parent.pendingNested.push(t);
			else out.push(t);
		}
	};
	let start = base;
	for (let cp = base; cp < end; cp++) {
		const ch = text[cp]!;
		const last = cp === end - 1;
		const isEnd = ch === '\r' || ch === '\x07' || (ch === '\x0c' && story === 'main' && sectionEnds.has(cp + 1));
		if (!isEnd && !last) continue;
		if (++counters.paragraphs > ctx.limits.maxParagraphs) throw new DocResourceLimitError('DOC paragraph count exceeds the processing limit');
		const fc = ctx.fcAt(cp);
		const mark = classify(ch, fc === undefined ? undefined : papx.at(fc));
		const pEnd = cp + 1;
		if (mark.level === 0) {
			closeTo(0);
		} else {
			closeTo(mark.level);
			while (stack.length < mark.level) {
				const parent = stack[stack.length - 1];
				if (parent && parent.cellStart === undefined) parent.cellStart = start;
				stack.push({ level: stack.length + 1, start: rel(start), end: rel(start), rows: [], row: undefined, cellStart: undefined, pendingNested: [] });
			}
			const t = stack[stack.length - 1]!;
			if (!t.row) t.row = { cells: [], start: rel(start) };
			if (t.cellStart === undefined) t.cellStart = start;
			if (mark.kind === 'cell') {
				if (++counters.cells > ctx.limits.maxCells) throw new DocResourceLimitError('DOC table cell count exceeds the processing limit');
				const nested = t.pendingNested;
				t.pendingNested = [];
				t.row.cells.push(makeCell(ctx, rel(t.cellStart), rel(cp), nested));
				t.cellStart = undefined;
			} else if (mark.kind === 'row') {
				t.rows.push(makeRow(t.row, rel(cp), rel(pEnd), mark.props));
				t.row = undefined;
				t.cellStart = undefined;
				t.end = rel(pEnd);
			}
		}
		start = pEnd;
	}
	closeTo(0);
	// A table can end only at a row mark; unfinished trailing rows are kept as incomplete rows.
	function makeCell(c: StoryContext, s: number, markCp: number, nested: DocTable[]): DocTableCell {
		return {
			start: s,
			end: markCp + 1,
			markCp,
			text: c.text.slice(base + s, base + markCp),
			gridColumn: 0,
			columnSpan: 1,
			rowSpan: 1,
			props: undefined,
			nestedTables: nested,
		};
	}
}

function makeRow(row: RowBuilder, markCp: number, end: number, p: DocParaProps | undefined): DocTableRow {
	const props = p ? readRowProps(p.tap) : undefined;
	return { start: row.start, end, markCp, cells: assignCells(row.cells, props), props };
}

/** Pair cell marks with TC entries; horizontally merged TCs have no mark of their own. */
function assignCells(marks: readonly DocTableCell[], props: DocRowProps | undefined): DocTableCell[] {
	if (!props) return marks.map((c, i) => ({ ...c, gridColumn: i }));
	const tcs = props.cells;
	const leading = tcs.map((c, i) => ({ c, i })).filter((x) => x.c.horzMerge !== 'continue');
	const byIndex = leading.length === marks.length;
	return marks.map((cell, k) => {
		const tcIndex = byIndex ? leading[k]!.i : k;
		let span = 1;
		if (byIndex) while (tcs[tcIndex + span]?.horzMerge === 'continue') span++;
		const tc = tcs[tcIndex];
		let width = tc?.widthTwips ?? 0;
		for (let s = 1; s < span; s++) width += tcs[tcIndex + s]!.widthTwips;
		return { ...cell, gridColumn: tcIndex, columnSpan: span, props: tc ? { ...tc, widthTwips: width } : undefined };
	});
}

function finish(story: DocStoryId, t: TableBuilder): DocTable {
	const rows = t.rows.slice();
	if (t.row && t.row.cells.length > 0) {
		// Unterminated trailing row: keep its cells rather than dropping content.
		const lastCell = t.row.cells[t.row.cells.length - 1]!;
		rows.push({ start: t.row.start, end: lastCell.end, markCp: lastCell.markCp, cells: assignCells(t.row.cells, undefined), props: undefined });
		t.end = lastCell.end;
	}
	const tableRows = applyGridColumns(rows);
	applyRowSpans(tableRows);
	return { story, level: t.level, start: t.start, end: t.end, rows: tableRows };
}

/** Map each cell onto the union of all rows' boundaries so merged and ragged rows share grid columns. */
function applyGridColumns(rows: DocTableRow[]): DocTableRow[] {
	const bounds: number[] = [];
	for (const r of rows) for (const g of r.props?.gridTwips ?? []) bounds.push(g);
	bounds.sort((a, b) => a - b);
	const grid: number[] = [];
	for (const b of bounds) if (grid.length === 0 || b - grid[grid.length - 1]! > 2) grid.push(b);
	const index = (v: number) => {
		let best = 0;
		for (let i = 1; i < grid.length; i++) if (Math.abs(grid[i]! - v) < Math.abs(grid[best]! - v)) best = i;
		return best;
	};
	return rows.map((row) => {
		const g = row.props?.gridTwips;
		if (!g) return row;
		return {
			...row,
			cells: row.cells.map((cell) => {
				const left = g[cell.gridColumn];
				const right = g[cell.gridColumn + cell.columnSpan];
				if (left === undefined || right === undefined) return cell;
				const from = index(left);
				return { ...cell, gridColumn: from, columnSpan: Math.max(1, index(right) - from) };
			}),
		};
	});
}

function applyRowSpans(rows: DocTableRow[]): void {
	for (let r = 0; r < rows.length; r++) {
		rows[r] = {
			...rows[r]!,
			cells: rows[r]!.cells.map((cell) => {
				if (cell.props?.vertMerge !== 'restart') return cell;
				let span = 1;
				for (let n = r + 1; n < rows.length; n++) {
					const below = rows[n]!.cells.find((c) => c.gridColumn === cell.gridColumn);
					if (below?.props?.vertMerge !== 'continue') break;
					span++;
				}
				return span > 1 ? { ...cell, rowSpan: span } : cell;
			}),
		};
	}
}
