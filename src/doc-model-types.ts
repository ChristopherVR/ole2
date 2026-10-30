/**
 * The composed Word 97-2003 document model returned by `readDocModel`: per
 * story an ordered block tree (paragraphs with formatted inline items, tables
 * of cells that hold blocks), plus sections, headers and footers, notes,
 * comments, text boxes, styles, lists, fonts and properties. It is shaped for
 * a consumer that maps it to WordprocessingML.
 *
 * Positions: every `cp*` field is an absolute document CP (see `doc-story`).
 *
 * @module doc-model-types
 */
import type { DocCharProps, DocFontNames } from './doc-char-props.js';
import type { DocFont, DocProperties } from './doc-fonts.js';
import type { DocParaProps } from './doc-para-props.js';
import type { DocListDefinition, DocListLevel, DocListOverride, DocLists } from './doc-lists.js';
import type { DocStyle } from './doc-styles.js';
import type { DocCellProps, DocRowProps } from './doc-table-props.js';
import type { DocSection } from './doc-sections.js';
import type { DocHeaderFooterKind, DocHeaderFooterVariant } from './doc-headers-footers.js';
import type { DocStoryId } from './doc-story.js';
import type {
	DocBlip, DocBookmark, DocField, DocFloatingShape, DocInlinePicture, DocOleObject,
} from './doc-objects-types.js';

/** Resolved character formatting shared between items; do not mutate. */
export type DocModelCharProps = Readonly<DocCharProps>;

/** A run of text with uniform formatting. Control characters are removed; tabs stay as `\t`. */
export interface DocTextItem {
	readonly type: 'text';
	readonly cpStart: number;
	readonly cpEnd: number;
	readonly text: string;
	readonly props: DocModelCharProps;
}

/** An explicit break: 0x0B line break, 0x0C page break or 0x0E column break. Section breaks end the paragraph instead. */
export interface DocBreakItem {
	readonly type: 'break';
	readonly kind: 'line' | 'page' | 'column';
	readonly cp: number;
	readonly props: DocModelCharProps;
}

/** A footnote or endnote separator mark (0x03 separator, 0x04 continuation separator) in a separator story. */
export interface DocSeparatorItem {
	readonly type: 'separator';
	readonly kind: 'separator' | 'continuationSeparator';
	readonly cp: number;
	readonly props: DocModelCharProps;
}

/** The auto-number mark at the start of a note or comment story (0x02 or 0x05). */
export interface DocNoteMarkItem {
	readonly type: 'noteMark';
	readonly kind: 'footnote' | 'endnote' | 'comment';
	readonly cp: number;
	readonly props: DocModelCharProps;
}

/** Inline picture (0x01 with a picture location). `picture.blip` carries the image bytes. */
export interface DocPictureItem {
	readonly type: 'picture';
	readonly picture: DocInlinePicture;
	readonly props: DocModelCharProps;
}

/** Embedded OLE object (0x01 with an object location); `object.preview` is its picture. */
export interface DocOleItem {
	readonly type: 'ole';
	readonly object: DocOleObject;
	readonly props: DocModelCharProps;
}

/** Complex field marks, in Word's own order: begin, instruction text, separator, result items, end. */
export interface DocFieldBeginItem {
	readonly type: 'fieldBegin';
	/** Index into {@link DocModel.fields}, or -1 when the field table has no entry for this mark. */
	readonly id: number;
	/** Typed field (HYPERLINK url, REF bookmark, ...), undefined when the field table has no entry. */
	readonly field: DocField | undefined;
	readonly cp: number;
	readonly props: DocModelCharProps;
}
export interface DocFieldInstructionItem {
	readonly type: 'fieldInstruction';
	readonly id: number;
	readonly text: string;
	readonly cpStart: number;
	readonly cpEnd: number;
	readonly props: DocModelCharProps;
}
export interface DocFieldSeparatorItem {
	readonly type: 'fieldSeparate';
	readonly id: number;
	readonly cp: number;
}
export interface DocFieldEndItem {
	readonly type: 'fieldEnd';
	readonly id: number;
	readonly cp: number;
}

export interface DocBookmarkStartItem {
	readonly type: 'bookmarkStart';
	/** Index into {@link DocModel.bookmarks}. */
	readonly id: number;
	readonly bookmark: DocBookmark;
}
export interface DocBookmarkEndItem {
	readonly type: 'bookmarkEnd';
	readonly id: number;
	readonly bookmark: DocBookmark;
}

/** Footnote or endnote reference in the main story; `id` indexes {@link DocModel.footnotes} or `endnotes`. */
export interface DocNoteReferenceItem {
	readonly type: 'noteReference';
	readonly kind: 'footnote' | 'endnote';
	readonly id: number;
	readonly autoNumbered: boolean;
	readonly cp: number;
	readonly props: DocModelCharProps;
}
/** Comment reference mark; `id` indexes {@link DocModel.comments}. */
export interface DocCommentReferenceItem {
	readonly type: 'commentReference';
	readonly id: number;
	readonly cp: number;
	readonly props: DocModelCharProps;
}

/** Anchor of a floating shape (0x08 drawn object) without text. */
export interface DocShapeAnchorItem {
	readonly type: 'shapeAnchor';
	readonly shape: DocFloatingShape;
	readonly props: DocModelCharProps;
}
/** Anchor of a floating text box; `textBox` indexes {@link DocModel.textBoxes}. */
export interface DocTextBoxAnchorItem {
	readonly type: 'textBoxAnchor';
	readonly shape: DocFloatingShape;
	readonly textBox: number;
	readonly props: DocModelCharProps;
}

export type DocInline =
	| DocTextItem | DocBreakItem | DocSeparatorItem | DocNoteMarkItem | DocPictureItem | DocOleItem
	| DocFieldBeginItem | DocFieldInstructionItem | DocFieldSeparatorItem | DocFieldEndItem
	| DocBookmarkStartItem | DocBookmarkEndItem | DocNoteReferenceItem | DocCommentReferenceItem
	| DocShapeAnchorItem | DocTextBoxAnchorItem;

/** The list a paragraph belongs to, resolved from `ilfo` and `ilvl`. */
export interface DocParagraphList {
	readonly ilfo: number;
	readonly level: number;
	readonly override: DocListOverride | undefined;
	readonly definition: DocListDefinition | undefined;
	/** The effective level formatting (an override level when present, else the definition's). */
	readonly levelDefinition: DocListLevel | undefined;
}

export interface DocParagraphBlock {
	readonly type: 'paragraph';
	readonly cpStart: number;
	/** Absolute CP just past the paragraph mark. */
	readonly cpEnd: number;
	/** Paragraph style index (istd) into {@link DocModel.styles}. */
	readonly style: number;
	/** Resolved paragraph properties (style cascade plus direct formatting). */
	readonly props: DocParaProps;
	/** Resolved character properties of the paragraph mark. */
	readonly markProps: DocModelCharProps;
	readonly list: DocParagraphList | undefined;
	readonly inlines: readonly DocInline[];
	/** Index into {@link DocModel.sections} of the section this paragraph ends (its mark carries the section break). */
	readonly endsSection: number | undefined;
}

export interface DocTableCellBlock {
	/** First grid column (0-based) of the cell. */
	readonly gridColumn: number;
	readonly columnSpan: number;
	/** Rows spanned when this cell restarts a vertical merge, else 1. */
	readonly rowSpan: number;
	/** `continue` cells are covered by the restarting cell above and carry no content of their own. */
	readonly vMerge: 'restart' | 'continue' | undefined;
	readonly props: DocCellProps | undefined;
	readonly blocks: readonly DocBlock[];
}

export interface DocTableRowBlock {
	readonly cpStart: number;
	readonly props: DocRowProps | undefined;
	readonly cells: readonly DocTableCellBlock[];
}

export interface DocTableBlock {
	readonly type: 'table';
	/** 1 for an outermost table, 2 and up for nested tables. */
	readonly level: number;
	readonly cpStart: number;
	readonly cpEnd: number;
	/** Grid column boundaries in twips (union over all rows). */
	readonly gridTwips: readonly number[];
	readonly rows: readonly DocTableRowBlock[];
}

export type DocBlock = DocParagraphBlock | DocTableBlock;

/** One story as a block tree. */
export interface DocModelStory {
	readonly story: DocStoryId;
	readonly cpStart: number;
	readonly cpEnd: number;
	readonly blocks: readonly DocBlock[];
}

/** Where a section finds the header or footer of one variant. */
export interface DocSectionHeaderFooterRef {
	/** Index into {@link DocModel.headersFooters} of the entry that supplies the content. */
	readonly index: number;
	/** The section has no story of its own and uses an earlier section's. */
	readonly inherited: boolean;
	/** The variant is active (first needs a title page, even needs facing pages). */
	readonly applies: boolean;
}

export interface DocModelSection {
	readonly section: DocSection;
	readonly headers: Partial<Record<DocHeaderFooterVariant, DocSectionHeaderFooterRef>>;
	readonly footers: Partial<Record<DocHeaderFooterVariant, DocSectionHeaderFooterRef>>;
}

/** One header or footer story slot of a section; `blocks` is empty when the section inherits. */
export interface DocModelHeaderFooter {
	readonly section: number;
	readonly kind: DocHeaderFooterKind;
	readonly variant: DocHeaderFooterVariant;
	readonly inherited: boolean;
	readonly applies: boolean;
	readonly story: DocModelStory;
}

export interface DocModelNote {
	/** Index within its list; the id used by `noteReference` items. */
	readonly id: number;
	readonly refCp: number;
	readonly autoNumbered: boolean;
	readonly story: DocModelStory;
}

export interface DocModelComment {
	readonly id: number;
	readonly refCp: number;
	readonly author: string | undefined;
	readonly initials: string | undefined;
	readonly date: string | undefined;
	readonly story: DocModelStory;
}

export interface DocModelTextBox {
	readonly id: number;
	readonly story: 'textboxes' | 'headerTextboxes';
	readonly shapeId: number | undefined;
	/** Anchor CP in the story the shape is anchored in (header text boxes: relative to the header story). */
	readonly anchorCp: number | undefined;
	readonly rectTwips: { left: number; top: number; right: number; bottom: number } | undefined;
	readonly content: DocModelStory;
}

export interface DocModelNoteSeparators {
	readonly footnoteSeparator: DocModelStory;
	readonly footnoteContinuationSeparator: DocModelStory;
	readonly footnoteContinuationNotice: DocModelStory;
	readonly endnoteSeparator: DocModelStory;
	readonly endnoteContinuationSeparator: DocModelStory;
	readonly endnoteContinuationNotice: DocModelStory;
}

/** The whole document. */
export interface DocModel {
	/** The main story. */
	readonly body: DocModelStory;
	readonly sections: readonly DocModelSection[];
	readonly headersFooters: readonly DocModelHeaderFooter[];
	readonly noteSeparators: DocModelNoteSeparators | undefined;
	readonly footnotes: readonly DocModelNote[];
	readonly endnotes: readonly DocModelNote[];
	readonly comments: readonly DocModelComment[];
	readonly textBoxes: readonly DocModelTextBox[];
	/** Every field, indexed by the ids of the field items. */
	readonly fields: readonly DocField[];
	/** Every bookmark, indexed by the ids of the bookmark items. */
	readonly bookmarks: readonly DocBookmark[];
	/** The document picture store. */
	readonly blips: readonly DocBlip[];
	/** Sparse array indexed by istd. */
	readonly styles: readonly (DocStyle | undefined)[];
	readonly lists: DocLists;
	readonly fonts: readonly DocFont[];
	/** Default font names per script class (document defaults, below the Normal style). */
	readonly defaultFonts: DocFontNames;
	readonly properties: DocProperties;
	/** Document-wide odd/even header distinction. */
	readonly facingPages: boolean;
	/** Recoverable problems; the affected part is omitted or approximated. */
	readonly warnings: readonly string[];
}
