/**
 * Structure of a legacy Word 97-2003 `.doc` beyond the main-body text:
 * sections, tables, headers and footers, footnotes, endnotes, comments and
 * text boxes, all addressed by character positions (CPs) so a caller can
 * compose it with a formatting reader. Read-only and bounded by
 * {@link DocStructureLimits}.
 *
 * @module doc-structure
 */
import { readHeadersFooters, type DocHeaderFooter, type DocNoteSeparators } from './doc-headers-footers.js';
import { readComments, readNotes, readTextBoxes, type DocComment, type DocNote, type DocTextBox } from './doc-notes.js';
import { readDocSections, readFacingPages, type DocSection } from './doc-sections.js';
import { openStoryContext, type DocStoryRange, type DocStructureLimits } from './doc-story.js';
import { readDocTables, type DocTable } from './doc-tables.js';

export interface DocStructure {
	readonly sections: readonly DocSection[];
	/** Outermost tables of every story, in story order; nested tables hang off their cells. */
	readonly tables: readonly DocTable[];
	readonly headersFooters: readonly DocHeaderFooter[];
	/** Footnote and endnote separator stories, when the document has a header story. */
	readonly noteSeparators: DocNoteSeparators | undefined;
	/** Document-wide odd/even header distinction (Dop.fFacingPages). */
	readonly facingPages: boolean;
	readonly footnotes: readonly DocNote[];
	readonly endnotes: readonly DocNote[];
	readonly comments: readonly DocComment[];
	readonly textboxes: readonly DocTextBox[];
	/** Main story as one range (CPs equal document-wide CPs). */
	readonly main: DocStoryRange;
	/** Recoverable problems; the affected part is omitted. */
	readonly warnings: readonly string[];
}

/**
 * Read the structure of a `.doc`. Throws for input that is not a readable,
 * unencrypted Word 97-2003 document and `DocResourceLimitError` when a budget
 * in `limits` is exceeded; damaged individual parts are skipped with a warning.
 */
export function readDocStructure(bytes: Uint8Array, limits?: DocStructureLimits): DocStructure {
	const ctx = openStoryContext(bytes, limits);
	const budget = { remaining: ctx.limits.maxSprms };
	const guard = <T>(name: string, fallback: T, fn: () => T): T => {
		try {
			return fn();
		} catch (e) {
			if (e instanceof Error && e.name === 'DocResourceLimitError') throw e;
			ctx.warnings.push(`${name}: ${e instanceof Error ? e.message : String(e)}`);
			return fallback;
		}
	};
	const sections = guard('sections', [] as DocSection[], () => readDocSections(ctx, budget));
	const facingPages = guard('dop', false, () => readFacingPages(ctx));
	const sectionEnds = new Set(sections.map((s) => s.end));
	const hf = guard('headers', { headersFooters: [] as DocHeaderFooter[], separators: undefined as DocNoteSeparators | undefined }, () =>
		readHeadersFooters(ctx, sections, facingPages),
	);
	return {
		sections,
		tables: guard('tables', [] as DocTable[], () => readDocTables(ctx, sectionEnds, budget)),
		headersFooters: hf.headersFooters,
		noteSeparators: hf.separators,
		facingPages,
		footnotes: guard('footnotes', [] as DocNote[], () => readNotes(ctx, 'footnotes')),
		endnotes: guard('endnotes', [] as DocNote[], () => readNotes(ctx, 'endnotes')),
		comments: guard('comments', [] as DocComment[], () => readComments(ctx)),
		textboxes: guard('textboxes', [] as DocTextBox[], () => readTextBoxes(ctx)),
		main: ctx.range('main', 0, ctx.extents.main.length),
		warnings: ctx.warnings,
	};
}
