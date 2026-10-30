import { readFileSync } from 'node:fs';
import { describe, expect, it } from 'vitest';
import { DocResourceLimitError, parseDocClx, parseGrpprl, readDocFib, readDocFormatting, tryWriteOleDocParagraphEdit, unwrapDocBytes } from '../src/index.js';
import type { DocFormatting, DocFormatParagraph, DocFormatRun } from '../src/index.js';

const load = (name: string): Uint8Array =>
	new Uint8Array(readFileSync(new URL(`./fixtures/doc/${name}.doc`, import.meta.url)));
const read = (name: string): DocFormatting => {
	const result = readDocFormatting(load(name));
	expect(result).toBeDefined();
	return result!;
};
const run = (p: DocFormatParagraph, text: string): DocFormatRun => {
	const found = p.runs.find((r) => r.text === text);
	expect(found, `run ${text}`).toBeDefined();
	return found!;
};

describe('readDocFormatting: character formatting (formatting-rich.doc)', () => {
	const doc = read('formatting-rich');
	const [title, words, boxed, hanging] = doc.paragraphs;

	it('splits the story into paragraphs with cp ranges that tile it', () => {
		expect(doc.paragraphs).toHaveLength(4);
		expect(doc.warnings).toEqual([]);
		let cp = 0;
		for (const p of doc.paragraphs) {
			expect(p.cpStart).toBe(cp);
			expect(p.terminator).toBe('\r');
			expect(p.cpEnd).toBe(p.cpStart + p.text.length + 1);
			expect(p.runs.map((r) => r.text).join('')).toBe(p.text);
			cp = p.cpEnd;
		}
		expect(cp).toBe(doc.characterCount);
		expect(title!.text).toBe('Centered title');
	});

	it('decodes toggles, underline, scripts, colour, highlight, size and spacing', () => {
		const w = words!;
		expect(run(w, 'bold').props.bold).toBe(true);
		expect(run(w, 'italic').props.italic).toBe(true);
		expect(run(w, 'underline').props.underline).toBe('double');
		expect(run(w, 'strike').props.strike).toBe(true);
		expect(run(w, 'dstrike').props.doubleStrike).toBe(true);
		expect(run(w, 'caps').props.caps).toBe(true);
		expect(run(w, 'small').props.smallCaps).toBe(true);
		expect(run(w, 'hidden').props.hidden).toBe(true);
		expect(run(w, 'super').props.verticalAlign).toBe('superscript');
		expect(run(w, 'sub').props.verticalAlign).toBe('subscript');
		expect(run(w, 'colour').props.color).toBe('C81020');
		expect(run(w, 'highlight').props.highlight).toBe('yellow');
		expect(run(w, 'big').props.sizePt).toBe(20);
		expect(run(w, 'spaced').props.spacingTwips).toBe(60);
		// Neighbouring text keeps the inherited 12 pt Arial and none of the direct flags.
		const plain = w.runs.find((r) => r.text === ' ')!;
		expect(plain.props.sizePt).toBe(12);
		expect(plain.props.bold).toBeUndefined();
		expect(plain.props.fonts?.ascii).toBe('Arial');
	});

	it('resolves fonts per script class through the font table', () => {
		const courier = run(words!, 'courier').props;
		expect(courier.fonts?.ascii).toBe('Courier New');
		expect(courier.fonts?.hAnsi).toBe('Courier New');
		expect(doc.fonts.map((f) => f.name)).toEqual(expect.arrayContaining(['Arial', 'Courier New', 'Times New Roman']));
	});

	it('decodes paragraph alignment, spacing and line rule', () => {
		expect(title!.props.alignment).toBe('center');
		expect(title!.props.spaceBeforeTwips).toBe(240);
		expect(title!.props.spaceAfterTwips).toBe(120);
		expect(title!.props.lineSpacing).toEqual({ value: 360, rule: 'auto' });
	});

	it('decodes indents, keep, page break, widow, outline, borders, shading and tabs', () => {
		const p = boxed!.props;
		expect(p.alignment).toBe('justify');
		expect(p.indentLeftTwips).toBe(720);
		expect(p.indentRightTwips).toBe(360);
		expect(p.firstLineTwips).toBe(360);
		expect(p.keepWithNext).toBe(true);
		expect(p.keepTogether).toBe(true);
		expect(p.pageBreakBefore).toBe(true);
		expect(p.widowControl).toBe(false);
		expect(p.outlineLevel).toBe(1);
		expect(p.borders?.top).toMatchObject({ style: 'single', widthPt: 1 });
		expect(p.borders?.bottom).toMatchObject({ style: 'single', widthPt: 0.5 });
		expect(p.borders?.left).toBeUndefined();
		expect(p.shading).toMatchObject({ background: 'FFFF00', fill: 'FFFF00' });
		expect(p.tabs).toEqual([
			{ positionTwips: 1440, alignment: 'center', leader: 'dot' },
			{ positionTwips: 4320, alignment: 'right', leader: 'none' },
		]);
	});

	it('decodes a hanging indent as a negative first-line indent', () => {
		expect(hanging!.props.indentLeftTwips).toBe(1080);
		expect(hanging!.props.firstLineTwips).toBe(-360);
		expect(hanging!.props.alignment).toBe('right');
	});

	it('reads the document default tab stop', () => {
		expect(doc.properties.defaultTabStopTwips).toBe(810);
		expect(doc.properties.facingPages).toBe(false);
	});
});

describe('readDocFormatting: styles (formatting-styles.doc)', () => {
	const doc = read('formatting-styles');
	const byName = (name: string) => doc.styles.find((s) => s?.name === name)!;

	it('reads names, types and the inheritance links of custom styles', () => {
		const body = byName('FixtureBody');
		const two = byName('FixtureBodyTwo');
		const emphasis = byName('FixtureEmphasis');
		expect(body).toMatchObject({ type: 'paragraph', builtIn: false, basedOn: 0 });
		expect(two).toMatchObject({ type: 'paragraph', basedOn: body.istd });
		expect(body.next).toBe(two.istd);
		expect(emphasis.type).toBe('character');
		expect(byName('Normal')).toMatchObject({ istd: 0, sti: 0, builtIn: true });
		expect(byName('Heading 1')).toMatchObject({ istd: 1, type: 'paragraph' });
	});

	it('keeps linked paragraph and character styles connected', () => {
		const linked = byName('FixtureBody Char');
		expect(linked.type).toBe('character');
		expect(linked.link).toBe(byName('FixtureBody').istd);
		expect(byName('FixtureBody').link).toBe(linked.istd);
	});

	it('exposes each style own formatting', () => {
		expect(byName('FixtureBody').paragraph).toMatchObject({ alignment: 'center', indentLeftTwips: 480, spaceAfterTwips: 200 });
		expect(byName('FixtureBody').character).toMatchObject({ bold: true, sizePt: 14, color: '0000A0' });
		expect(byName('FixtureEmphasis').character).toMatchObject({ italic: true, underline: 'single', color: '008000' });
	});

	it('cascades styles into paragraphs and keeps the raw style references', () => {
		const [heading, first, second, normal] = doc.paragraphs;
		expect(heading!.style).toBe(1);
		expect(heading!.runs[0]!.props).toMatchObject({ sizePt: 20, color: '800000' });
		expect(heading!.runs[0]!.props.fonts?.ascii).toBe('Georgia');
		expect(heading!.props.outlineLevel).toBe(0);
		expect(first!.style).toBe(byName('FixtureBody').istd);
		expect(first!.props).toMatchObject({ alignment: 'center', indentLeftTwips: 480 });
		expect(first!.runs[0]!.props).toMatchObject({ bold: true, sizePt: 14 });
		expect(first!.runs[0]!.props.fonts?.ascii).toBe('Verdana');
		// The derived style inverts the inherited bold, adds italic and inherits everything else.
		expect(second!.style).toBe(byName('FixtureBodyTwo').istd);
		expect(second!.runs[0]!.props).toMatchObject({ bold: false, italic: true, sizePt: 14, color: '0000A0' });
		expect(second!.props.alignment).toBe('center');
		expect(normal!.style).toBe(0);
	});

	it('applies a character style over the paragraph style and records its istd', () => {
		const normal = doc.paragraphs[3]!;
		const styled = run(normal, 'styled');
		expect(styled.style).toBe(byName('FixtureEmphasis').istd);
		expect(styled.props).toMatchObject({ italic: true, underline: 'single', color: '008000', sizePt: 12 });
		expect(normal.runs.filter((r) => r.style === undefined)).toHaveLength(2);
	});
});

describe('readDocFormatting: lists (formatting-lists.doc)', () => {
	const doc = read('formatting-lists');
	const def = (ilfo: number) => {
		const override = doc.lists.overrides[ilfo - 1]!;
		return doc.lists.definitions[override.definitionIndex!]!;
	};

	it('links paragraphs to list instances and levels', () => {
		const refs = doc.paragraphs.map((p) => [p.props.ilfo ?? 0, p.props.ilvl ?? 0]);
		expect(refs).toEqual([[1, 0], [1, 0], [2, 0], [2, 0], [3, 0], [3, 1], [3, 2], [3, 0], [0, 0]]);
		expect(doc.lists.overrides).toHaveLength(3);
		expect(doc.lists.overrides.every((o) => o.definitionIndex !== undefined)).toBe(true);
	});

	it('reads a bullet list with its bullet font and indents', () => {
		const level = def(1).levels[0]!;
		expect(level).toMatchObject({ isBullet: true, numberFormat: 23, numberFormatName: 'bullet', follow: 'tab' });
		expect(level.text).toHaveLength(1);
		expect(level.text.charCodeAt(0)).toBe(0xf0b7);
		expect(level.character.fonts?.ascii).toBe('Symbol');
		expect(level.paragraph).toMatchObject({ indentLeftTwips: 720, firstLineTwips: -360 });
	});

	it('reads a lettered list with start value and level text', () => {
		const list = def(2);
		expect(list.simple).toBe(true);
		expect(list.levels[0]).toMatchObject({ numberFormatName: 'lowerLetter', text: '%1)', startAt: 3, isBullet: false });
		expect(list.levels[0]!.paragraph).toMatchObject({ indentLeftTwips: 600, firstLineTwips: -600 });
	});

	it('reads all nine levels of a multilevel list with placeholders', () => {
		const list = def(3);
		expect(list.simple).toBe(false);
		expect(list.levels).toHaveLength(9);
		expect(list.levels.slice(0, 3).map((l) => l.text)).toEqual(['%1.', '%1.%2', '(%3)']);
		expect(list.levels.slice(0, 3).map((l) => l.numberFormatName)).toEqual(['decimal', 'decimal', 'lowerRoman']);
		expect(list.levels.slice(0, 3).map((l) => l.paragraph.indentLeftTwips)).toEqual([720, 1440, 2160]);
		expect(list.levels.map((l) => l.level)).toEqual([0, 1, 2, 3, 4, 5, 6, 7, 8]);
	});

	it('resolves list indents of the paragraph style cascade', () => {
		// The List Paragraph style supplies the left indent and contextual spacing.
		expect(doc.paragraphs[0]!.props).toMatchObject({ indentLeftTwips: 720, contextualSpacing: true });
	});
});

describe('readDocFormatting: Unicode piece and appended pieces (formatting-unicode.doc)', () => {
	const doc = read('formatting-unicode');

	it('stores the text in a single Unicode piece', () => {
		const cfb = unwrapDocBytes(load('formatting-unicode'))!;
		const pieces = parseDocClx(cfb.tableBytes, readDocFib(cfb.wordDocBytes).clx).pieces;
		expect(pieces.map((p) => p.compressed)).toEqual([false]);
	});

	it('reads a multi-piece file produced by the piece-append editor', () => {
		const edited = tryWriteOleDocParagraphEdit(load('formatting-rich'), 1, 'bold replaced 中 text');
		expect(edited).toMatchObject({ status: 'edited', strategy: 'piece-append' });
		const cfb = unwrapDocBytes(edited.bytes)!;
		const pieces = parseDocClx(cfb.tableBytes, readDocFib(cfb.wordDocBytes).clx).pieces;
		expect(pieces.length).toBeGreaterThan(2);
		expect(pieces.some((p) => p.compressed)).toBe(true);
		expect(pieces.some((p) => !p.compressed)).toBe(true);
		const doc = readDocFormatting(edited.bytes)!;
		expect(doc.paragraphs.map((p) => p.text)).toEqual([
			'Centered title',
			'bold replaced 中 text',
			'Indented justified paragraph with borders shading and tabs.',
			'Hanging indent paragraph.',
		]);
		for (const p of doc.paragraphs) expect(p.runs.map((r) => r.text).join('')).toBe(p.text);
		expect(doc.paragraphs.at(-1)!.cpEnd).toBe(doc.characterCount);
		// Untouched paragraphs keep their formatting across the pieces.
		expect(doc.paragraphs[0]!.props.alignment).toBe('center');
		expect(doc.paragraphs[2]!.props).toMatchObject({ alignment: 'justify', indentLeftTwips: 720 });
		expect(doc.paragraphs[3]!.props.firstLineTwips).toBe(-360);
	});

	it('reads text across several pieces and keeps cp ranges contiguous', () => {
		expect(doc.paragraphs.map((p) => p.text)).toEqual([
			'First fast paragraph.',
			'Inserted Second fast paragraph.',
			'Third fast paragraph.Appended é中 fast text.',
			'',
		]);
		expect(doc.paragraphs.at(-1)!.cpEnd).toBe(doc.characterCount);
	});

	it('keeps formatting applied after the first save and Unicode piece fonts', () => {
		expect(run(doc.paragraphs[0]!, 'First').props.bold).toBe(true);
		expect(doc.paragraphs[0]!.runs[1]!.props.bold).toBeUndefined();
		const cjk = run(doc.paragraphs[2]!, '中');
		expect(cjk.props.fonts?.eastAsia).toBe('MS Gothic');
		expect(doc.paragraphs[2]!.runs.map((r) => r.text).join('')).toBe(doc.paragraphs[2]!.text);
	});
});

describe('readDocFormatting: robustness', () => {
	it('returns undefined for bytes that are not a Word document', () => {
		expect(readDocFormatting(new Uint8Array(0))).toBeUndefined();
		expect(readDocFormatting(new Uint8Array(512))).toBeUndefined();
		expect(readDocFormatting(new Uint8Array(readFileSync(new URL('./fixtures/sample-deck.ppt', import.meta.url))))).toBeUndefined();
	});

	it('rejects invalid limits and enforces processing budgets', () => {
		const bytes = load('formatting-rich');
		expect(() => readDocFormatting(bytes, { maxParagraphs: -1 })).toThrow(/nonnegative/);
		expect(() => readDocFormatting(bytes, { maxParagraphs: 2 })).toThrow(DocResourceLimitError);
		expect(() => readDocFormatting(bytes, { maxRuns: 3 })).toThrow(DocResourceLimitError);
		expect(() => readDocFormatting(bytes, { maxCharacters: 10 })).toThrow(DocResourceLimitError);
		expect(() => readDocFormatting(bytes, { maxSprms: 5 })).toThrow(DocResourceLimitError);
		expect(() => readDocFormatting(bytes, { maxFormattingPages: 0 })).toThrow(DocResourceLimitError);
	});

	it('does not throw on truncated or corrupted files', () => {
		const bytes = load('formatting-rich');
		for (const cut of [600, 2048, 8192, 16384]) {
			expect(() => readDocFormatting(bytes.subarray(0, cut))).not.toThrow();
		}
		for (let seed = 1; seed <= 25; seed++) {
			const copy = bytes.slice();
			let state = seed * 2654435761;
			for (let i = 0; i < 60; i++) {
				state = (state * 1664525 + 1013904223) >>> 0;
				copy[state % copy.length] = state >>> 24;
			}
			try {
				readDocFormatting(copy);
			} catch (e) {
				expect(e).toBeInstanceOf(DocResourceLimitError);
			}
		}
	});
});

describe('parseGrpprl', () => {
	const budget = () => ({ remaining: 100 });

	it('decodes fixed and length-prefixed operands', () => {
		const sprms = parseGrpprl(new Uint8Array([0x35, 0x08, 0x01, 0x43, 0x4a, 0x28, 0x00, 0x4e, 0xc6, 0x02, 0xaa, 0xbb]), budget());
		expect(sprms.map((s) => [s.opcode, [...s.operand]])).toEqual([
			[0x0835, [1]],
			[0x4a43, [0x28, 0]],
			[0xc64e, [0xaa, 0xbb]],
		]);
	});

	it('stops at a truncated record and honours the budget', () => {
		expect(parseGrpprl(new Uint8Array([0x35, 0x08, 0x01, 0x43, 0x4a, 0x28]), budget())).toHaveLength(1);
		expect(() => parseGrpprl(new Uint8Array([0x35, 0x08, 0x01, 0x35, 0x08, 0x00]), { remaining: 1 })).toThrow(DocResourceLimitError);
	});

	it('sizes the table definition and oversized tab operands', () => {
		const tdef = parseGrpprl(new Uint8Array([0x08, 0xd6, 0x04, 0x00, 1, 2, 3, 0x35, 0x08, 0x01]), budget());
		expect(tdef.map((s) => s.operand.length)).toEqual([3, 1]);
		// cb == 255: one deleted tab (position and close, 4 bytes), one added tab (position and descriptor, 3 bytes).
		const tabs = parseGrpprl(new Uint8Array([0x15, 0xc6, 0xff, 1, 0x0a, 0, 0x0b, 0, 1, 0x0c, 0, 0, 0x35, 0x08, 0x01]), budget());
		expect(tabs.map((s) => s.operand.length)).toEqual([9, 1]);
	});
});
