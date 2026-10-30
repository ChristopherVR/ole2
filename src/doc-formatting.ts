/**
 * Rich legacy `.doc` formatting reader: stylesheet, lists, fonts, and every
 * story (main text, headers and footers, footnotes, endnotes, comments, text
 * boxes) split into paragraphs and runs with resolved character and paragraph
 * properties ([MS-DOC] 2.4 and 2.6). Everything is keyed by absolute character
 * position (CP) so table, field and picture readers compose on top of the same
 * ranges.
 *
 * @module doc-formatting
 */
import type { DocCharProps, DocFontNames } from './doc-char-props.js';
import { readDocFonts, readDocProperties } from './doc-fonts.js';
import type { DocFont, DocProperties } from './doc-fonts.js';
import { splitParagraphs } from './doc-formatting-paragraphs.js';
import type { SplitCounters } from './doc-formatting-paragraphs.js';
import { resolveDocLimits } from './doc-limits.js';
import type { DocModelLimits } from './doc-limits.js';
import { readDocLists } from './doc-lists.js';
import type { DocLists } from './doc-lists.js';
import type { DocParaProps } from './doc-para-props.js';
import { readDocSections } from './doc-sections.js';
import { openStoryContext, STORY_ORDER } from './doc-story.js';
import type { DocStoryId, StoryContext } from './doc-story.js';
import { DocStyleResolver } from './doc-style-resolve.js';
import { readDocStylesheet } from './doc-styles.js';
import type { DocStyle } from './doc-styles.js';
import { DocResourceLimitError } from './ole-document-doc-pieces.js';

/** A run: a CP range with uniform resolved character formatting. */
export interface DocFormatRun {
	/** Absolute CP of the first character. */
	cpStart: number;
	cpEnd: number;
	text: string;
	/** Character style (istd) applied directly to the run, if any. */
	style: number | undefined;
	/** Fully resolved properties (defaults, paragraph style, character style, direct). Shared between runs: do not mutate. */
	props: Readonly<DocCharProps>;
}

/** A paragraph: `[cpStart, cpEnd)` (absolute CPs) includes its terminating mark; `text` excludes it. */
export interface DocFormatParagraph {
	cpStart: number;
	cpEnd: number;
	text: string;
	/** Paragraph style (istd). */
	style: number;
	/** Resolved paragraph properties; style cascade plus direct formatting. */
	props: DocParaProps;
	/** Resolved character properties of the paragraph mark (what an empty paragraph looks like). */
	markProps: Readonly<DocCharProps>;
	/** The terminating character: carriage return, the table cell/row mark (U+0007), or a section form feed (U+000C). */
	terminator: string;
	runs: DocFormatRun[];
}

/** Everything {@link readDocFormatting} extracts. */
export interface DocFormatting {
	/** Sparse array indexed by istd. */
	styles: (DocStyle | undefined)[];
	lists: DocLists;
	fonts: DocFont[];
	/** Default font names per script class (the stylesheet header's ftc defaults). */
	defaultFonts: DocFontNames;
	properties: DocProperties;
	/** Main story length in characters (ccpText). */
	characterCount: number;
	/** Paragraphs of the main story (the same array as `stories.main`). */
	paragraphs: DocFormatParagraph[];
	/** Paragraphs of every non-empty story, with absolute CPs. */
	stories: Partial<Record<DocStoryId, DocFormatParagraph[]>>;
	/** Non-fatal problems (unreadable stylesheet, list table, ...). */
	warnings: string[];
}

/** Processing budgets (a subset of {@link DocModelLimits}). */
export type DocFormattingLimits = DocModelLimits;

const FIB_STSH = 1, FIB_FFN = 15, FIB_DOP = 31, FIB_LST = 73, FIB_LFO = 74;

/**
 * Read styles, lists, fonts and every formatted story from an opened context.
 * Damaged optional tables become warnings; budget overruns throw
 * {@link DocResourceLimitError}.
 */
export function formatFromContext(ctx: StoryContext, sectionEnds: ReadonlySet<number>): DocFormatting {
	const warnings: string[] = [];
	const attempt = <T>(label: string, fallback: T, fn: () => T): T => {
		try {
			return fn();
		} catch (e) {
			if (e instanceof DocResourceLimitError) throw e;
			warnings.push(`${label} unreadable: ${e instanceof Error ? e.message : String(e)}`);
			return fallback;
		}
	};
	const { table, budget, limits } = ctx;
	const fonts = attempt('font table', [] as DocFont[], () => readDocFonts(table, ctx.pair(FIB_FFN)));
	const sheet = attempt('stylesheet', { styles: [], defaultFonts: {}, stiMaxWhenSaved: 0 }, () =>
		readDocStylesheet(table, ctx.pair(FIB_STSH), { maxStyles: limits.maxStyles, budget }));
	const resolver = new DocStyleResolver(sheet, fonts);
	const lists = attempt('lists', { definitions: [], overrides: [] } as DocLists, () =>
		readDocLists(table, ctx.pair(FIB_LST), ctx.pair(FIB_LFO), {
			maxLists: limits.maxLists, budget, fontNames: (p) => resolver.withFontNames(p),
		}));
	const properties = attempt('document properties', {} as DocProperties, () => readDocProperties(table, ctx.pair(FIB_DOP)));
	const counters: SplitCounters = { paragraphs: 0, runs: 0 };
	const stories: Partial<Record<DocStoryId, DocFormatParagraph[]>> = {};
	for (const id of STORY_ORDER) {
		const e = ctx.extents[id];
		if (e.length <= 0) continue;
		stories[id] = splitParagraphs({ ctx, resolver, cpStart: e.start, cpEnd: e.start + e.length, sectionEnds, counters });
	}
	const defaultFonts = resolver.withFontNames({ fontIds: { ...sheet.defaultFonts } }).fonts ?? {};
	return {
		styles: sheet.styles, lists, fonts, defaultFonts, properties, characterCount: ctx.extents.main.length,
		paragraphs: stories.main ?? [], stories, warnings: [...ctx.warnings, ...warnings],
	};
}

/**
 * Read styles, lists, fonts and the formatted stories of a legacy Word
 * 97-2003 document. Returns undefined when the bytes are not a readable,
 * unencrypted Word document. Throws {@link DocResourceLimitError} when a
 * processing budget is exceeded. Control characters stay in the text (field
 * marks, picture anchors, line and page breaks) so other readers can
 * interpret them by CP.
 */
export function readDocFormatting(bytes: Uint8Array, limits?: DocFormattingLimits): DocFormatting | undefined {
	resolveDocLimits(limits); // invalid limits are a caller error, not a damaged file
	try {
		const ctx = openStoryContext(bytes, limits);
		if (ctx.text.length < ctx.extents.main.length) return undefined;
		let sectionEnds = new Set<number>();
		try {
			sectionEnds = new Set(readDocSections(ctx).map((s) => s.end));
		} catch (e) {
			if (e instanceof DocResourceLimitError) throw e;
		}
		return formatFromContext(ctx, sectionEnds);
	} catch (e) {
		if (e instanceof DocResourceLimitError) throw e;
		return undefined;
	}
}
