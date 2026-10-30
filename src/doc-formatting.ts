/**
 * Rich legacy `.doc` formatting reader: stylesheet, lists, fonts, and the
 * main story split into paragraphs and runs with resolved character and
 * paragraph properties ([MS-DOC] 2.4 and 2.6). Everything is keyed by
 * character position (CP) so table, field and picture readers compose on
 * top of the same ranges.
 *
 * @module doc-formatting
 */
import type { DocCharProps } from './doc-char-props.js';
import { readDocFonts, readDocProperties } from './doc-fonts.js';
import type { DocFont, DocProperties } from './doc-fonts.js';
import { readFkpRuns } from './doc-fkp-runs.js';
import { readDocLists } from './doc-lists.js';
import type { DocLists } from './doc-lists.js';
import type { DocParaProps } from './doc-para-props.js';
import { DocStyleResolver } from './doc-style-resolve.js';
import { readDocStylesheet } from './doc-styles.js';
import type { DocStyle } from './doc-styles.js';
import type { SprmBudget } from './doc-sprm.js';
import { splitParagraphs } from './doc-formatting-paragraphs.js';
import { unwrapDocBytes } from './ole-document-doc-cfb.js';
import { DEFAULT_OLE_DOC_PROCESSING_LIMITS } from './ole-document-doc-editor.js';
import type { OleDocProcessingLimits } from './ole-document-doc-editor.js';
import { readDocFib, readFcLcbAt } from './ole-document-doc-fib.js';
import type { ParsedDocFib } from './ole-document-doc-fib.js';
import { DocResourceLimitError, decodePiecesText, parseDocClx } from './ole-document-doc-pieces.js';

/** A run: a CP range with uniform resolved character formatting. */
export interface DocFormatRun {
	cpStart: number;
	cpEnd: number;
	text: string;
	/** Character style (istd) applied directly to the run, if any. */
	style: number | undefined;
	/** Fully resolved properties (defaults, paragraph style, character style, direct). Shared between runs: do not mutate. */
	props: Readonly<DocCharProps>;
}

/** A paragraph: `[cpStart, cpEnd)` includes its terminating mark; `text` excludes it. */
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
	/** The terminating character: carriage return, or the table cell/row mark (U+0007). */
	terminator: string;
	runs: DocFormatRun[];
}

/** Everything {@link readDocFormatting} extracts. */
export interface DocFormatting {
	/** Sparse array indexed by istd. */
	styles: (DocStyle | undefined)[];
	lists: DocLists;
	fonts: DocFont[];
	properties: DocProperties;
	/** Main story length in characters (ccpText). */
	characterCount: number;
	paragraphs: DocFormatParagraph[];
	/** Non-fatal problems (unreadable stylesheet, list table, ...). */
	warnings: string[];
}

/** Processing budgets; the inherited limits keep their {@link DEFAULT_OLE_DOC_PROCESSING_LIMITS} defaults. */
export interface DocFormattingLimits extends OleDocProcessingLimits {
	/** Paragraphs materialised (default 262144). */
	maxParagraphs?: number;
	/** Runs materialised (default 1048576). */
	maxRuns?: number;
	/** Styles read (default 4096). */
	maxStyles?: number;
	/** Sprms decoded in total (default 8388608). */
	maxSprms?: number;
	/** Formatted disk pages read per table (default 65536). */
	maxFormattingPages?: number;
	/** List definitions and instances read (default 4096). */
	maxLists?: number;
}

const FIB_STSH = 1, FIB_FFN = 15, FIB_DOP = 31, FIB_LST = 73, FIB_LFO = 74;
const NONE = { fc: 0, lcb: 0 };

function limitOf(value: number | undefined, fallback: number): number {
	if (value === undefined) return fallback;
	if (!Number.isSafeInteger(value) || value < 0) throw new Error('DOC processing limits must be nonnegative safe integers');
	return value;
}

function pair(word: Uint8Array, fib: ParsedDocFib, index: number): { fc: number; lcb: number } {
	try {
		return readFcLcbAt(word, fib, index);
	} catch {
		return NONE;
	}
}

/**
 * Read styles, lists, fonts and the formatted main story of a legacy Word
 * 97-2003 document. Returns undefined when the bytes are not a readable,
 * unencrypted Word document. Throws {@link DocResourceLimitError} when a
 * processing budget is exceeded. Control characters stay in the text (field
 * marks, picture anchors, line and page breaks) so other readers can
 * interpret them by CP.
 */
export function readDocFormatting(bytes: Uint8Array, limits?: DocFormattingLimits): DocFormatting | undefined {
	const maxCharacters = limitOf(limits?.maxCharacters, DEFAULT_OLE_DOC_PROCESSING_LIMITS.maxCharacters);
	const maxPieces = limitOf(limits?.maxPieces, DEFAULT_OLE_DOC_PROCESSING_LIMITS.maxPieces);
	const budget: SprmBudget = { remaining: limitOf(limits?.maxSprms, 8_388_608) };
	const maxParagraphs = limitOf(limits?.maxParagraphs, 262_144);
	const maxRuns = limitOf(limits?.maxRuns, 1_048_576);
	const maxStyles = limitOf(limits?.maxStyles, 4096);
	const maxLists = limitOf(limits?.maxLists, 4096);
	const maxPages = limitOf(limits?.maxFormattingPages, 65_536);
	try {
		const doc = unwrapDocBytes(bytes);
		if (!doc) return undefined;
		const word = doc.wordDocBytes, table = doc.tableBytes;
		const fib = readDocFib(word);
		if (fib.ccpText > maxCharacters) throw new DocResourceLimitError('DOC main story exceeds the processing limit');
		const clx = parseDocClx(table, fib.clx, maxPieces);
		const text = decodePiecesText(word, clx.pieces, fib.ccpText);
		if (text.length < fib.ccpText) return undefined;
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
		const fonts = attempt('font table', [] as DocFont[], () => readDocFonts(table, pair(word, fib, FIB_FFN)));
		const sheet = attempt('stylesheet', { styles: [], defaultFonts: {}, stiMaxWhenSaved: 0 }, () =>
			readDocStylesheet(table, pair(word, fib, FIB_STSH), { maxStyles, budget }));
		const resolver = new DocStyleResolver(sheet, fonts);
		const lists = attempt('lists', { definitions: [], overrides: [] } as DocLists, () =>
			readDocLists(table, pair(word, fib, FIB_LST), pair(word, fib, FIB_LFO), {
				maxLists, budget, fontNames: (p) => resolver.withFontNames(p),
			}));
		const properties = attempt('document properties', {} as DocProperties, () => readDocProperties(table, pair(word, fib, FIB_DOP)));
		const limitsFkp = { maxPages, maxRuns };
		const chpx = readFkpRuns(word, table, fib.plcfbteChpx, 'chpx', limitsFkp);
		const papx = readFkpRuns(word, table, fib.plcfbtePapx, 'papx', limitsFkp);
		const paragraphs = splitParagraphs({
			text: text.slice(0, fib.ccpText), pieces: clx.pieces, prcBytes: clx.prcBytes, chpx, papx, resolver, budget,
			maxParagraphs, maxRuns,
		});
		return { styles: sheet.styles, lists, fonts, properties, characterCount: fib.ccpText, paragraphs, warnings };
	} catch (e) {
		if (e instanceof DocResourceLimitError) throw e;
		return undefined;
	}
}
