import { readFileSync, readdirSync } from 'node:fs';
import { describe, expect, it } from 'vitest';
import { DocResourceLimitError, readDocFormatting, readDocModel } from '../src/index.js';
import type { DocBlock, DocInline, DocModel, DocParagraphBlock, DocTableBlock } from '../src/index.js';

// Fixtures are generated through Word COM (test/fixtures/doc/generate-*.ps1); expected values are what those scripts set.
const bytes = (name: string) => new Uint8Array(readFileSync(new URL(`./fixtures/doc/${name}.doc`, import.meta.url)));
const model = (name: string): DocModel => {
	const m = readDocModel(bytes(name));
	expect(m, name).toBeDefined();
	return m!;
};
const paragraphs = (blocks: readonly DocBlock[]): DocParagraphBlock[] => blocks.filter((b): b is DocParagraphBlock => b.type === 'paragraph');
const tables = (blocks: readonly DocBlock[]): DocTableBlock[] => blocks.filter((b): b is DocTableBlock => b.type === 'table');
const text = (p: DocParagraphBlock): string => p.inlines.map((i) => (i.type === 'text' ? i.text : '')).join('');
const types = (p: DocParagraphBlock): string[] => p.inlines.map((i) => i.type);
const allInlines = (blocks: readonly DocBlock[]): DocInline[] =>
	blocks.flatMap((b) => (b.type === 'paragraph' ? [...b.inlines] : b.rows.flatMap((r) => r.cells.flatMap((c) => allInlines(c.blocks)))));
const cellTexts = (t: DocTableBlock) => t.rows.map((r) => r.cells.map((c) => paragraphs(c.blocks).map(text).join('|')));

describe('readDocModel: every fixture', () => {
	const names = readdirSync(new URL('./fixtures/doc/', import.meta.url)).filter((f) => f.endsWith('.doc')).map((f) => f.slice(0, -4));
	it.each(names)('composes %s with no control characters in text and consistent CPs', (name) => {
		const m = model(name);
		expect(m.body.story).toBe('main');
		expect(m.sections.length).toBeGreaterThan(0);
		const stories = [m.body, ...m.headersFooters.map((h) => h.story), ...m.footnotes.map((n) => n.story), ...m.endnotes.map((n) => n.story),
			...m.comments.map((c) => c.story), ...m.textBoxes.map((t) => t.content)];
		for (const story of stories) {
			let last = story.cpStart;
			for (const i of allInlines(story.blocks)) {
				if (i.type === 'text' || i.type === 'fieldInstruction') {
					expect(i.text).not.toMatch(/[\u0000-\u0008\u000a-\u001f]/u);
					expect(i.cpStart).toBeGreaterThanOrEqual(story.cpStart);
					expect(i.cpEnd).toBeLessThanOrEqual(story.cpEnd);
					expect(i.cpEnd - i.cpStart).toBe(i.text.length);
					if (i.type === 'text' && story === m.body) {
						expect(i.cpStart).toBeGreaterThanOrEqual(last);
						last = i.cpEnd;
					}
				}
			}
		}
		// Every field begin has a matching end, in order.
		const depth = allInlines(m.body.blocks).reduce((d, i) => d + (i.type === 'fieldBegin' ? 1 : i.type === 'fieldEnd' ? -1 : 0), 0);
		expect(depth).toBe(0);
	});
});

describe('readDocModel: formatting fixtures', () => {
	it('keeps list references on list paragraphs (formatting-lists)', () => {
		const m = model('formatting-lists');
		const ps = paragraphs(m.body.blocks);
		expect(ps.map(text)).toEqual(['Bullet one', 'Bullet two', 'Number one', 'Number two', 'Level one', 'Level two', 'Level three', 'Level one again', 'Plain closing paragraph.']);
		expect(ps.map((p) => p.list && [p.list.ilfo, p.list.level])).toEqual([[1, 0], [1, 0], [2, 0], [2, 0], [3, 0], [3, 1], [3, 2], [3, 0], undefined]);
		expect(ps[0]!.list!.levelDefinition!.isBullet).toBe(true);
		expect(ps[2]!.list!.levelDefinition!.isBullet).toBe(false);
		expect(ps[2]!.list!.definition).toBe(m.lists.definitions[ps[2]!.list!.override!.definitionIndex!]);
		expect(ps[5]!.list!.levelDefinition!.level).toBe(1);
	});

	it('reports the default fonts of the stylesheet header', () => {
		const m = model('structure-notes');
		expect(m.defaultFonts.ascii).toBeTruthy();
		expect(m.defaultFonts.cs).toBe('Times New Roman');
	});

	it('formats runs and paragraphs (formatting-rich)', () => {
		const m = model('formatting-rich');
		const ps = paragraphs(m.body.blocks);
		expect(ps[0]!.props.alignment).toBe('center');
		const runs = ps[1]!.inlines.filter((i) => i.type === 'text');
		const by = (t: string) => runs.find((r) => r.type === 'text' && r.text === t);
		expect(by('bold')!.type === 'text' && by('bold')!.props.bold).toBe(true);
		expect(by('italic')!.type === 'text' && by('italic')!.props.italic).toBe(true);
		expect(by('courier')!.type === 'text' && by('courier')!.props.fonts?.ascii).toBe('Courier New');
		expect(ps[2]!.props.borders?.top).toMatchObject({ style: 'single' });
		expect(ps[2]!.props.shading?.fill).toBe('FFFF00');
		expect(ps[3]!.props.firstLineTwips).toBe(-360);
	});

	it('resolves styles (formatting-styles)', () => {
		const m = model('formatting-styles');
		const ps = paragraphs(m.body.blocks);
		expect(m.styles[ps[0]!.style]?.name).toBe('Heading 1');
		expect(m.styles[ps[1]!.style]?.type).toBe('paragraph');
		const styled = ps[3]!.inlines.filter((i) => i.type === 'text');
		expect(styled.map((i) => (i.type === 'text' ? i.text : ''))).toEqual(['Normal paragraph with a ', 'styled', ' word inside.']);
		expect(styled.some((i) => i.type === 'text' && i.props.styleId !== undefined)).toBe(true);
	});

	it('reads fast-saved Unicode text (formatting-unicode)', () => {
		const m = model('formatting-unicode');
		const joined = paragraphs(m.body.blocks).map(text).join('\n');
		expect(joined).toContain('Appended é');
		expect(joined).toContain('中');
	});

	it('keeps plain run formatting (paragraph-alignment, rich-runs, rich-size-runs, underline-runs)', () => {
		for (const name of ['paragraph-alignment', 'rich-runs', 'rich-size-runs', 'underline-runs']) {
			const ps = paragraphs(model(name).body.blocks);
			expect(ps.map(text), name).toEqual(['First plain paragraph.', 'Bold text, italic text, plain text.', 'Last plain paragraph.']);
			const bold = ps[1]!.inlines.find((i) => i.type === 'text' && i.text === 'Bold text');
			expect(bold?.type === 'text' && bold.props.bold, name).toBe(true);
		}
	});

	it('formats every story, not only the main one', () => {
		const f = readDocFormatting(bytes('header-field'))!;
		expect(Object.keys(f.stories).sort()).toEqual(['headers', 'main']);
		expect(f.stories.headers!.map((p) => p.text)).toContain('Fixture header \u0013 PAGE  \\* MERGEFORMAT \u00141\u0015');
		const n = readDocFormatting(bytes('structure-notes'))!;
		expect(Object.keys(n.stories).sort()).toEqual(['comments', 'endnotes', 'footnotes', 'headers', 'main']);
		expect(n.stories.footnotes![0]!.cpStart).toBeGreaterThan(n.paragraphs.at(-1)!.cpEnd - 1);
		expect(n.stories.comments![0]!.runs[0]!.props).toBeDefined();
	});
});

describe('readDocModel: fields and objects (objects.doc, header-field.doc, main-field.doc)', () => {
	const m = model('objects');
	const ps = paragraphs(m.body.blocks);

	it('wraps a hyperlink field around its result text', () => {
		const p = ps.find((x) => text(x).startsWith('Link: '))!;
		const begin = p.inlines.findIndex((i) => i.type === 'fieldBegin' && i.field?.kind === 'hyperlink');
		const first = p.inlines[begin]!;
		expect(first.type === 'fieldBegin' && first.field).toMatchObject({ kind: 'hyperlink', url: 'https://example.com/path?x=1', tooltip: 'Example tip' });
		const rest = p.inlines.slice(begin + 1);
		expect(rest.map((i) => i.type).slice(0, 4)).toEqual(['fieldInstruction', 'fieldSeparate', 'text', 'fieldEnd']);
		const id = first.type === 'fieldBegin' ? first.id : -1;
		expect(rest[0]).toMatchObject({ type: 'fieldInstruction', id });
		expect(rest[2]).toMatchObject({ type: 'text', text: 'ext link' });
		expect(rest[3]).toMatchObject({ type: 'fieldEnd', id });
		expect(m.fields[id]!.kind).toBe('hyperlink');
	});

	it('places inline pictures with their blips', () => {
		const pics = allInlines(m.body.blocks).filter((i) => i.type === 'picture');
		expect(pics.length).toBeGreaterThanOrEqual(2);
		expect(pics[0]!.type === 'picture' && pics[0]!.picture.blip?.format).toBe('png');
		expect(pics[1]!.type === 'picture' && pics[1]!.picture.blip?.format).toBe('jpeg');
		const first = ps[0]!;
		expect(types(first)).toEqual(expect.arrayContaining(['bookmarkStart', 'shapeAnchor', 'picture', 'bookmarkEnd']));
		expect(text(first)).toBe('Inline PNG:  and inline JPEG:  end.');
	});

	it('places bookmarks around their range', () => {
		expect(m.bookmarks.map((b) => b.name)).toEqual(expect.arrayContaining(['BM_Target', 'BM_Point']));
		const items = allInlines(m.body.blocks).filter((i) => i.type === 'bookmarkStart' || i.type === 'bookmarkEnd');
		for (const [id, b] of m.bookmarks.entries()) {
			const start = items.filter((i) => i.type === 'bookmarkStart' && i.id === id);
			const end = items.filter((i) => i.type === 'bookmarkEnd' && i.id === id);
			expect([b.name, start.length, end.length]).toEqual([b.name, 1, 1]);
		}
	});

	it('places the embedded object once, after its field separator', () => {
		const oles = allInlines(m.body.blocks).filter((i) => i.type === 'ole');
		expect(oles).toHaveLength(1);
		expect(oles[0]!.type === 'ole' && oles[0]!.object).toMatchObject({ progId: 'Excel.Sheet.8' });
		const p = ps.find((x) => text(x).startsWith('Embedded: '))!;
		const at = p.inlines.findIndex((i) => i.type === 'ole');
		expect(p.inlines[at - 1]!.type === 'fieldSeparate' || p.inlines[at - 1]!.type === 'bookmarkEnd').toBe(true);
	});

	it('anchors the floating text box and keeps a field spanning paragraphs balanced', () => {
		const anchor = allInlines(m.body.blocks).find((i) => i.type === 'textBoxAnchor');
		expect(anchor?.type === 'textBoxAnchor' && m.textBoxes[anchor.textBox]!.shapeId).toBe(anchor?.type === 'textBoxAnchor' ? anchor.shape.spid : -1);
		expect(paragraphs(m.textBoxes[0]!.content.blocks).map(text)).toEqual(['Textbox content']);
		const toc = ps.findIndex((p) => p.inlines.some((i) => i.type === 'fieldBegin' && i.field?.kind === 'toc'));
		expect(ps[toc]!.inlines.some((i) => i.type === 'fieldEnd' && i.id === m.fields.findIndex((f) => f.kind === 'toc'))).toBe(false);
		expect(ps[toc + 1]!.inlines[0]).toMatchObject({ type: 'fieldEnd' });
	});

	it('keeps the header PAGE field and the main QUOTE field', () => {
		const h = model('header-field');
		const def = h.headersFooters.find((x) => x.kind === 'header' && x.variant === 'default')!;
		const inl = allInlines(def.story.blocks);
		expect(inl.filter((i) => i.type === 'fieldInstruction').map((i) => (i.type === 'fieldInstruction' ? i.text : '')).join('')).toContain('PAGE');
		expect(inl.find((i) => i.type === 'fieldBegin')).toMatchObject({ field: { kind: 'page' } });
		expect(h.sections[0]!.headers.default).toMatchObject({ index: h.headersFooters.indexOf(def), inherited: false, applies: true });
		const q = allInlines(model('main-field').body.blocks);
		expect(q.find((i) => i.type === 'fieldBegin')?.type).toBe('fieldBegin');
		expect(q.filter((i) => i.type === 'text').map((i) => (i.type === 'text' ? i.text : ''))).toContain('Fixture field');
	});
});

describe('readDocModel: structure fixtures', () => {
	it('builds a simple table with row and cell properties', () => {
		const m = model('structure-table-simple');
		const blocks = m.body.blocks;
		expect(blocks.map((b) => b.type)).toEqual(['paragraph', 'table', 'paragraph', 'paragraph']);
		const t = tables(blocks)[0]!;
		expect(t.level).toBe(1);
		expect(cellTexts(t)).toEqual([['R1C1', 'R1C2'], ['R2C1', 'R2C2'], ['R3C1', 'R3C2']]);
		expect(t.rows[0]!.props?.borders.top).toMatchObject({ style: 'single', widthPt: 0.5 });
		expect(t.rows[1]!.cells[0]!.props?.shading?.background).toBe('FF0000');
		expect(t.gridTwips.length).toBe(3);
		expect(text(paragraphs(blocks)[0]!)).toBe('Before table');
		expect(text(paragraphs(blocks).at(-1)!)).toBe('After table');
	});

	it('builds merged cells with spans', () => {
		const t = tables(model('structure-table-merged').body.blocks)[0]!;
		const r1 = t.rows[0]!.cells;
		expect(r1.map((c) => [c.gridColumn, c.columnSpan])).toEqual([[0, 2], [2, 1]]);
		expect(paragraphs(r1[0]!.blocks).map(text)).toEqual(['R1C1', 'R1C2']);
		const merged = t.rows[1]!.cells[2]!;
		expect(merged).toMatchObject({ rowSpan: 2, vMerge: 'restart' });
		expect(paragraphs(merged.blocks).map(text)).toEqual(['R2C3', 'R3C3']);
		expect(t.rows[2]!.cells[2]).toMatchObject({ vMerge: 'continue' });
	});

	it('nests a table inside a cell, between its paragraphs', () => {
		const t = tables(model('structure-table-nested').body.blocks)[0]!;
		const host = t.rows[0]!.cells[1]!;
		expect(host.blocks.map((b) => b.type)).toEqual(['paragraph', 'table', 'paragraph']);
		const inner = tables(host.blocks)[0]!;
		expect(inner.level).toBe(2);
		expect(cellTexts(inner)).toEqual([['In11', 'In12'], ['In21', 'In22']]);
		expect(cellTexts(t)[1]).toEqual(['Outer21', 'Outer22']);
	});

	it('ends each section on its last paragraph and links headers and footers', () => {
		const s = model('structure-sections');
		expect(paragraphs(s.body.blocks).map((p) => [text(p), p.endsSection])).toEqual([
			['Section one', 0], ['Section two', 1], ['Section three', 2], ['Section four', 3],
		]);
		expect(s.sections.map((x) => x.section.index)).toEqual([0, 1, 2, 3]);
		const h = model('structure-headers');
		expect(h.sections).toHaveLength(2);
		const own = h.sections[0]!.headers.default!;
		expect(text(paragraphs(h.headersFooters[own.index]!.story.blocks)[0]!)).toBe('Odd header 1');
		expect(h.sections[0]!.headers.first).toMatchObject({ applies: true });
		expect(h.sections[0]!.headers.even).toMatchObject({ applies: true });
		expect(text(paragraphs(h.headersFooters[h.sections[1]!.headers.default!.index]!.story.blocks)[0]!)).toBe('Odd header 2');
		const inherited = h.sections[1]!.footers.default!;
		expect(inherited.inherited).toBe(true);
		expect(text(paragraphs(h.headersFooters[inherited.index]!.story.blocks)[0]!)).toBe('Odd footer 1');
		// the extra paragraph mark that ends every header story is not content
		expect(paragraphs(h.headersFooters[own.index]!.story.blocks)).toHaveLength(1);
	});

	it('references footnotes, endnotes and comments by id', () => {
		const m = model('structure-notes');
		const items = allInlines(m.body.blocks);
		const refs = items.filter((i) => i.type === 'noteReference' || i.type === 'commentReference');
		expect(refs.map((r) => [r.type, r.type === 'noteReference' ? r.kind : 'comment', r.id])).toEqual([
			['commentReference', 'comment', 0], ['noteReference', 'footnote', 0], ['commentReference', 'comment', 1],
			['noteReference', 'footnote', 1], ['noteReference', 'endnote', 0], ['noteReference', 'endnote', 1],
		]);
		expect(m.footnotes.map((n) => text(paragraphs(n.story.blocks)[0]!).trim())).toEqual(['First footnote text', 'Second footnote text']);
		expect(m.endnotes.map((n) => text(paragraphs(n.story.blocks)[0]!).trim())).toEqual(['First endnote text', 'Second endnote text']);
		expect(paragraphs(m.footnotes[0]!.story.blocks)[0]!.inlines[0]).toMatchObject({ type: 'noteMark', kind: 'footnote' });
		expect(m.comments.map((c) => [c.author, c.initials, text(paragraphs(c.story.blocks)[0]!)])).toEqual([
			['Alice Author', expect.any(String), 'First comment text'], ['Bob Builder', expect.any(String), 'Second comment text'],
		]);
		expect(m.noteSeparators!.footnoteSeparator.blocks).toHaveLength(1);
		expect(paragraphs(m.noteSeparators!.footnoteSeparator.blocks)[0]!.inlines[0]).toMatchObject({ type: 'separator', kind: 'separator' });
	});

	it('anchors text boxes to their content (structure-textbox)', () => {
		const m = model('structure-textbox');
		const anchors = allInlines(m.body.blocks).filter((i) => i.type === 'textBoxAnchor');
		expect(anchors).toHaveLength(2);
		expect(anchors.map((a) => (a.type === 'textBoxAnchor' ? text(paragraphs(m.textBoxes[a.textBox]!.content.blocks)[0]!) : ''))).toEqual(['Text box one', 'Text box two']);
		expect(text(paragraphs(m.body.blocks)[0]!)).toBe('Body text with boxes.');
	});
});

describe('readDocModel: limits and hostile input', () => {
	it('returns undefined for input that is not a Word document', () => {
		expect(readDocModel(new Uint8Array(0))).toBeUndefined();
		expect(readDocModel(new Uint8Array(1000))).toBeUndefined();
		expect(readDocModel(new Uint8Array(readFileSync(new URL('./fixtures/sample-deck.ppt', import.meta.url))))).toBeUndefined();
	});

	it('rejects invalid limits and enforces budgets', () => {
		const t = bytes('structure-table-nested');
		expect(() => readDocModel(t, { maxParagraphs: -1 })).toThrow(/nonnegative/);
		expect(() => readDocModel(t, { maxCharacters: 10 })).toThrow(DocResourceLimitError);
		expect(() => readDocModel(t, { maxParagraphs: 3 })).toThrow(DocResourceLimitError);
		expect(() => readDocModel(t, { maxRuns: 2 })).toThrow(DocResourceLimitError);
		expect(() => readDocModel(t, { maxCells: 2 })).toThrow(DocResourceLimitError);
		expect(() => readDocModel(t, { maxSprms: 5 })).toThrow(DocResourceLimitError);
		expect(() => readDocModel(t, { maxNesting: 1 })).toThrow(DocResourceLimitError);
		expect(() => readDocModel(bytes('objects'), { maxRecords: 1 })).toThrow(DocResourceLimitError);
		expect(readDocModel(t, { maxNesting: 3 })).toBeDefined();
	});

	it('never throws anything but a resource error on corrupted copies of every fixture', () => {
		const names = readdirSync(new URL('./fixtures/doc/', import.meta.url)).filter((f) => f.endsWith('.doc')).map((f) => f.slice(0, -4));
		let seed = 0x2545f491;
		const next = () => ((seed ^= seed << 13), (seed ^= seed >>> 17), (seed ^= seed << 5), seed >>> 0);
		let defined = 0;
		for (const name of names) {
			const original = bytes(name);
			for (let round = 0; round < 12; round++) {
				const copy = original.slice();
				const flips = 1 + (next() % 24);
				for (let i = 0; i < flips; i++) copy[next() % copy.length] = next() & 0xff;
				const input = round % 4 === 3 ? copy.subarray(0, next() % copy.length) : copy;
				try {
					if (readDocModel(input, { maxCharacters: 1 << 20, maxParagraphs: 20_000, maxRuns: 50_000, maxSprms: 500_000 })) defined++;
				} catch (e) {
					expect(e, `${name} round ${round}`).toBeInstanceOf(DocResourceLimitError);
				}
			}
		}
		expect(defined).toBeGreaterThan(0);
	});
});
