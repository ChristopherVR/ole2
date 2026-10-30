/**
 * `readDocModel`: compose the formatting, object and structure readers over one
 * shared context into a single block-tree model of a legacy Word 97-2003
 * document, ready to be mapped to WordprocessingML. See `docs/doc-model.md`.
 *
 * @module doc-model
 */
import { formatFromContext } from './doc-formatting.js';
import type { DocFormatParagraph } from './doc-formatting.js';
import { resolveDocLimits } from './doc-limits.js';
import type { DocModelLimits } from './doc-limits.js';
import { buildRangeBlocks } from './doc-model-blocks.js';
import type { BlockEnv } from './doc-model-blocks.js';
import { bookmarkItemMap, newFieldState } from './doc-model-inline.js';
import type { InlineEnv } from './doc-model-inline.js';
import type {
	DocModel, DocModelComment, DocModelHeaderFooter, DocModelNote, DocModelNoteSeparators, DocModelSection,
	DocModelStory, DocModelTextBox, DocSectionHeaderFooterRef,
} from './doc-model-types.js';
import { objectsFromContext } from './doc-objects.js';
import { structureFromContext } from './doc-structure.js';
import type { DocStructure } from './doc-structure.js';
import { openStoryContext } from './doc-story.js';
import type { DocStoryId, DocStoryRange, StoryContext } from './doc-story.js';
import { DocResourceLimitError } from './ole-document-doc-pieces.js';

const EXTRA_END_MARK: ReadonlySet<DocStoryId> = new Set(['headers', 'textboxes', 'headerTextboxes']);

/**
 * Read a legacy `.doc` into the composed model. Returns undefined when the
 * bytes are not a readable, unencrypted Word 97-2003 document. Throws
 * {@link DocResourceLimitError} when a budget in `limits` is exceeded and
 * `Error` for invalid limits; damage inside one optional structure is skipped
 * and listed in `warnings`.
 */
export function readDocModel(bytes: Uint8Array, limits?: DocModelLimits): DocModel | undefined {
	resolveDocLimits(limits); // invalid limits are a caller error, not a damaged file
	try {
		const ctx = openStoryContext(bytes, limits);
		if (ctx.text.length < ctx.extents.main.length) return undefined;
		return composeModel(ctx);
	} catch (e) {
		if (e instanceof DocResourceLimitError) throw e;
		return undefined;
	}
}

function composeModel(ctx: StoryContext): DocModel {
	const structure = structureFromContext(ctx);
	const sectionEnds = new Map<number, number>();
	structure.sections.forEach((s, i) => sectionEnds.set(s.end, i));
	const formatting = formatFromContext(ctx, new Set(sectionEnds.keys()));
	const objects = objectsFromContext(ctx);

	const textBoxIds = new Map<number, number>();
	structure.textboxes.forEach((t, i) => {
		if (t.shapeId !== undefined && !textBoxIds.has(t.shapeId)) textBoxIds.set(t.shapeId, i);
	});
	const fieldIds = new Map<number, number>();
	objects.fields.forEach((f, i) => fieldIds.set(f.cpBegin, i));
	const noteRefs = (notes: DocStructure['footnotes']) => new Map(notes.map((n) => [n.refCp, { id: n.index, autoNumbered: n.autoNumbered }]));
	const inline: InlineEnv = {
		ctx,
		pictures: new Map(objects.inlinePictures.map((p) => [p.cp, p])),
		oles: new Map(objects.oleObjects.map((o) => [o.cp, o])),
		shapes: new Map(objects.floatingShapes.map((s) => [s.cp, s])),
		fieldIds,
		fields: objects.fields,
		footnoteRefs: noteRefs(structure.footnotes),
		endnoteRefs: noteRefs(structure.endnotes),
		commentRefs: new Map(structure.comments.map((c) => [c.refCp, c.index])),
		textBoxIds,
		bookmarkItems: bookmarkItemMap(objects.bookmarks),
	};
	const env: BlockEnv = { inline, lists: formatting.lists, sectionEnds, storyBase: (s) => ctx.extents[s].start };
	const tablesOf = (story: DocStoryId) => structure.tables.filter((t) => t.story === story);
	const tables = new Map<DocStoryId, ReturnType<typeof tablesOf>>();
	const storyTables = (story: DocStoryId) => {
		let t = tables.get(story);
		if (!t) tables.set(story, (t = tablesOf(story)));
		return t;
	};
	const paragraphs = (story: DocStoryId): readonly DocFormatParagraph[] => formatting.stories[story] ?? [];
	const slice = (range: DocStoryRange): DocModelStory => {
		const blocks = range.cpEnd > range.cpStart
			? [...buildRangeBlocks(env, paragraphs(range.story), range.cpStart, range.cpEnd, storyTables(range.story), newFieldState())]
			: [];
		// Header, separator and text box stories end with one extra bare paragraph mark that is not content.
		const end = blocks.at(-1);
		if (EXTRA_END_MARK.has(range.story) && blocks.length > 1 && end?.type === 'paragraph' && end.inlines.length === 0) blocks.pop();
		return { story: range.story, cpStart: range.cpStart, cpEnd: range.cpEnd, blocks };
	};

	const body = slice(structure.main);
	const headersFooters: DocModelHeaderFooter[] = structure.headersFooters.map((h) => ({
		section: h.section, kind: h.kind, variant: h.variant, inherited: h.inherited, applies: h.applies, story: slice(h.range),
	}));
	const slot = new Map(structure.headersFooters.map((h, i) => [`${h.section}:${h.kind}:${h.variant}`, i]));
	const sections: DocModelSection[] = structure.sections.map((section, k) => {
		const headers: Partial<Record<string, DocSectionHeaderFooterRef>> = {};
		const footers: Partial<Record<string, DocSectionHeaderFooterRef>> = {};
		for (const h of structure.headersFooters) {
			if (h.section !== k || h.sourceSection === undefined) continue;
			const index = slot.get(`${h.sourceSection}:${h.kind}:${h.variant}`);
			if (index === undefined) continue;
			(h.kind === 'header' ? headers : footers)[h.variant] = { index, inherited: h.inherited, applies: h.applies };
		}
		return { section, headers, footers };
	});
	const sep = structure.noteSeparators;
	const noteSeparators: DocModelNoteSeparators | undefined = sep && {
		footnoteSeparator: slice(sep.footnoteSeparator),
		footnoteContinuationSeparator: slice(sep.footnoteContinuationSeparator),
		footnoteContinuationNotice: slice(sep.footnoteContinuationNotice),
		endnoteSeparator: slice(sep.endnoteSeparator),
		endnoteContinuationSeparator: slice(sep.endnoteContinuationSeparator),
		endnoteContinuationNotice: slice(sep.endnoteContinuationNotice),
	};
	const note = (n: DocStructure['footnotes'][number]): DocModelNote => ({ id: n.index, refCp: n.refCp, autoNumbered: n.autoNumbered, story: slice(n.range) });
	const comments: DocModelComment[] = structure.comments.map((c) => ({
		id: c.index, refCp: c.refCp, author: c.author, initials: c.initials, date: c.date, story: slice(c.range),
	}));
	const textBoxes: DocModelTextBox[] = structure.textboxes.map((t, i) => ({
		id: i, story: t.story, shapeId: t.shapeId, anchorCp: t.anchorCp, rectTwips: t.rectTwips, content: slice(t.range),
	}));

	return {
		body,
		sections,
		headersFooters,
		noteSeparators,
		footnotes: structure.footnotes.map(note),
		endnotes: structure.endnotes.map(note),
		comments,
		textBoxes,
		fields: objects.fields,
		bookmarks: objects.bookmarks,
		blips: objects.blips,
		styles: formatting.styles,
		lists: formatting.lists,
		fonts: formatting.fonts,
		properties: formatting.properties,
		facingPages: structure.facingPages,
		warnings: [...new Set([...ctx.warnings, ...formatting.warnings, ...objects.warnings, ...structure.warnings])],
	};
}
