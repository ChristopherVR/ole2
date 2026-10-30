/**
 * Turn one formatted paragraph into ordered inline items: formatted text with
 * control characters stripped, explicit breaks, and the pictures, field marks,
 * bookmarks, note and comment references and shape anchors placed at their CPs.
 *
 * @module doc-model-inline
 */
import type { DocFormatParagraph } from './doc-formatting.js';
import type { DocBookmark, DocField, DocFloatingShape, DocInlinePicture, DocOleObject } from './doc-objects-types.js';
import type { StoryContext } from './doc-story.js';
import type { DocInline, DocModelCharProps } from './doc-model-types.js';

/** Lookups from absolute CP to the object placed there, shared by every story of one read. */
export interface InlineEnv {
	readonly ctx: StoryContext;
	readonly pictures: ReadonlyMap<number, DocInlinePicture>;
	readonly oles: ReadonlyMap<number, DocOleObject>;
	readonly shapes: ReadonlyMap<number, DocFloatingShape>;
	/** Absolute begin-mark CP to field id (index into the field list). */
	readonly fieldIds: ReadonlyMap<number, number>;
	readonly fields: readonly DocField[];
	readonly footnoteRefs: ReadonlyMap<number, { id: number; autoNumbered: boolean }>;
	readonly endnoteRefs: ReadonlyMap<number, { id: number; autoNumbered: boolean }>;
	readonly commentRefs: ReadonlyMap<number, number>;
	/** Shape id to text box id. */
	readonly textBoxIds: ReadonlyMap<number, number>;
	/** Bookmark start and end items by the CP they sit before. */
	readonly bookmarkItems: ReadonlyMap<number, readonly DocInline[]>;
}

/** Field nesting state of one story; fields may span paragraphs, so it outlives a single paragraph. */
export interface FieldState {
	readonly stack: { id: number; region: 'instruction' | 'result' }[];
}

export function newFieldState(): FieldState {
	return { stack: [] };
}

/** Index bookmarks into start/end items keyed by CP; ends sort before starts at the same CP. */
export function bookmarkItemMap(bookmarks: readonly DocBookmark[]): Map<number, DocInline[]> {
	const keyed: { cp: number; order: number; item: DocInline }[] = [];
	bookmarks.forEach((bookmark, id) => {
		keyed.push({ cp: bookmark.cpStart, order: 1, item: { type: 'bookmarkStart', id, bookmark } });
		keyed.push({ cp: bookmark.cpEnd, order: bookmark.cpEnd === bookmark.cpStart ? 2 : 0, item: { type: 'bookmarkEnd', id, bookmark } });
	});
	keyed.sort((a, b) => a.cp - b.cp || a.order - b.order);
	const out = new Map<number, DocInline[]>();
	for (const k of keyed) {
		const list = out.get(k.cp);
		if (list) list.push(k.item);
		else out.set(k.cp, [k.item]);
	}
	return out;
}

/** Inline items of a paragraph, in CP order. */
export function buildInlines(env: InlineEnv, p: DocFormatParagraph, state: FieldState): DocInline[] {
	const out: DocInline[] = [];
	const text = env.ctx.text;
	const story = env.ctx.storyOf(p.cpStart);
	const markers = env.bookmarkItems;
	const stack = state.stack;
	let buf = '';
	let bufStart = 0;
	let props: DocModelCharProps = p.markProps;

	const flush = (cp: number): void => {
		if (!buf) return;
		const top = stack[stack.length - 1];
		if (top && top.region === 'instruction') {
			out.push({ type: 'fieldInstruction', id: top.id, text: buf, cpStart: bufStart, cpEnd: cp, props });
		} else {
			out.push({ type: 'text', cpStart: bufStart, cpEnd: cp, text: buf, props });
		}
		buf = '';
	};
	const place = (cp: number): void => {
		const items = markers.get(cp);
		if (items) {
			flush(cp);
			out.push(...items);
		}
	};

	for (const run of p.runs) {
		props = run.props;
		for (let cp = run.cpStart; cp < run.cpEnd; cp++) {
			if (markers.size > 0) place(cp);
			const ch = text.charCodeAt(cp);
			if (ch >= 0x20 || ch === 9) {
				if (!buf) bufStart = cp;
				buf += text[cp]!;
				continue;
			}
			flush(cp);
			special(ch, cp);
		}
		flush(run.cpEnd);
	}
	if (markers.size > 0 && p.terminator) place(p.cpEnd - 1);
	return out;

	function oleAt(cp: number): void {
		const ole = env.oles.get(cp);
		if (ole) out.push({ type: 'ole', object: ole, props });
	}

	function special(ch: number, cp: number): void {
		switch (ch) {
			case 0x01: {
				const ole = env.oles.get(cp);
				const picture = env.pictures.get(cp);
				if (ole) out.push({ type: 'ole', object: ole, props });
				else if (picture) out.push({ type: 'picture', picture, props });
				break;
			}
			case 0x02: {
				if (story === 'footnotes' || story === 'endnotes') {
					out.push({ type: 'noteMark', kind: story === 'footnotes' ? 'footnote' : 'endnote', cp, props });
					break;
				}
				const foot = env.footnoteRefs.get(cp);
				const end = foot ? undefined : env.endnoteRefs.get(cp);
				const ref = foot ?? end;
				if (ref) out.push({ type: 'noteReference', kind: foot ? 'footnote' : 'endnote', id: ref.id, autoNumbered: ref.autoNumbered, cp, props });
				break;
			}
			case 0x03:
			case 0x04:
				out.push({ type: 'separator', kind: ch === 3 ? 'separator' : 'continuationSeparator', cp, props });
				break;
			case 0x05: {
				if (story === 'comments') {
					out.push({ type: 'noteMark', kind: 'comment', cp, props });
					break;
				}
				const id = env.commentRefs.get(cp);
				if (id !== undefined) out.push({ type: 'commentReference', id, cp, props });
				break;
			}
			case 0x08: {
				const shape = env.shapes.get(cp);
				if (!shape) break;
				const textBox = shape.textboxId ? env.textBoxIds.get(shape.spid) : undefined;
				if (textBox !== undefined) out.push({ type: 'textBoxAnchor', shape, textBox, props });
				else out.push({ type: 'shapeAnchor', shape, props });
				break;
			}
			case 0x0b:
				out.push({ type: 'break', kind: 'line', cp, props });
				break;
			case 0x0c:
				out.push({ type: 'break', kind: 'page', cp, props });
				break;
			case 0x0e:
				out.push({ type: 'break', kind: 'column', cp, props });
				break;
			case 0x13: {
				const id = env.fieldIds.get(cp) ?? -1;
				out.push({ type: 'fieldBegin', id, field: id >= 0 ? env.fields[id] : undefined, cp, props });
				stack.push({ id, region: 'instruction' });
				oleAt(cp);
				break;
			}
			case 0x14: {
				const top = stack[stack.length - 1];
				if (!top) break;
				top.region = 'result';
				out.push({ type: 'fieldSeparate', id: top.id, cp });
				oleAt(cp); // Word stores the OLE object properties on the separator; its preview picture follows
				break;
			}
			case 0x15: {
				const top = stack.pop();
				if (top) out.push({ type: 'fieldEnd', id: top.id, cp });
				break;
			}
			case 0x1e:
			case 0x1f:
				if (!buf) bufStart = cp;
				buf += ch === 0x1e ? '‑' : '­';
				break;
			default:
		}
	}
}
