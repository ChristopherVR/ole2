/**
 * Processing budgets shared by every `.doc` reader (formatting, objects,
 * structure and the composed model). None is a format validity limit; each is
 * checked before the allocation it guards.
 *
 * @module doc-limits
 */
import type { OleDocProcessingLimits } from './ole-document-doc-editor.js';

/** Every budget the `.doc` readers honour. A reader uses the subset it needs. */
export interface DocModelLimits extends OleDocProcessingLimits {
	/** Paragraphs materialised or scanned across all stories (default 1,000,000). */
	maxParagraphs?: number;
	/** Formatting runs materialised (default 1,048,576). */
	maxRuns?: number;
	/** Styles read (default 4,096). */
	maxStyles?: number;
	/** Property modifiers decoded in total (default 8,388,608). */
	maxSprms?: number;
	/** Formatted disk pages read per table (default 65,536). */
	maxFormattingPages?: number;
	/** List definitions and instances read (default 4,096). */
	maxLists?: number;
	/** Table cells produced (default 262,144). */
	maxCells?: number;
	/** Entries of any one PLC (sections, notes, comments, text boxes, header stories; default 65,536). */
	maxEntries?: number;
	/** Records per table (fields, bookmarks, shapes, BLIPs, pictures, objects; default 65,536). */
	maxRecords?: number;
	/** Largest single image, compressed or decompressed (default 67,108,864). */
	maxBlipBytes?: number;
	/** Total image bytes returned (default 268,435,456). */
	maxTotalBlipBytes?: number;
	/** OfficeArt container nesting depth (default 16). */
	maxArtDepth?: number;
	/** Nesting depth of tables and fields in the block tree (default 32). */
	maxNesting?: number;
}

export type ResolvedDocLimits = Readonly<Required<DocModelLimits>>;

export const DEFAULT_DOC_LIMITS: ResolvedDocLimits = Object.freeze({
	maxCharacters: 16_777_216,
	maxPieces: 65_536,
	maxFieldRecords: 65_536,
	maxParagraphs: 1_000_000,
	maxRuns: 1_048_576,
	maxStyles: 4096,
	maxSprms: 8_388_608,
	maxFormattingPages: 65_536,
	maxLists: 4096,
	maxCells: 262_144,
	maxEntries: 65_536,
	maxRecords: 65_536,
	maxBlipBytes: 67_108_864,
	maxTotalBlipBytes: 268_435_456,
	maxArtDepth: 16,
	maxNesting: 32,
});

/** Merge `limits` over the defaults; throws for negative or non-integer values. */
export function resolveDocLimits(limits?: DocModelLimits): ResolvedDocLimits {
	const out = {
		...DEFAULT_DOC_LIMITS,
		...Object.fromEntries(Object.entries(limits ?? {}).filter(([, v]) => v !== undefined)),
	} as Required<DocModelLimits>;
	for (const v of Object.values(out)) {
		if (!Number.isSafeInteger(v) || v < 0) throw new Error('DOC processing limits must be nonnegative safe integers');
	}
	return Object.freeze(out);
}
