import { readFileSync } from 'node:fs';
import { describe, expect, it } from 'vitest';
import { readOleDocParagraphs, writeOleDocParagraphEdit } from '../src/ole-document-doc-editor.js';

const fixture = new URL('./fixtures/ole-word-97.doc', import.meta.url);
const paragraphs = [
	'First paragraph plain text.',
	'Second paragraph has a bold word in the middle.',
	'Third paragraph, plain again, this is the one we will edit.',
	'Fourth and final paragraph.',
];
const loadFixture = () => new Uint8Array(readFileSync(fixture));

describe('Word 97-2003 piece table editing', () => {
	it('reads main-body paragraph text from the real fixture', () => {
		expect(readOleDocParagraphs(loadFixture())).toStrictEqual(paragraphs);
	});

	it('replaces paragraphs in any position and preserves the other paragraph text', () => {
		for (const index of [0, 2, 3]) {
			const updated = writeOleDocParagraphEdit(loadFixture(), index, `Replacement at ${index}.`);
			const expected = [...paragraphs];
			expected[index] = `Replacement at ${index}.`;
			expect(readOleDocParagraphs(updated)).toStrictEqual(expected);
		}
	});

	it('handles growth, Windows-1252 and UTF-16 text across repeated edits', () => {
		let updated = writeOleDocParagraphEdit(loadFixture(), 1, 'Café naïve, curly ‘quotes’. ' + 'Long text. '.repeat(100));
		updated = writeOleDocParagraphEdit(updated, 3, 'Emoji fallback 😀 end.');
		expect(readOleDocParagraphs(updated)).toStrictEqual([
			paragraphs[0], 'Café naïve, curly ‘quotes’. ' + 'Long text. '.repeat(100), paragraphs[2], 'Emoji fallback 😀 end.',
		]);
	});

	it('normalizes embedded paragraph breaks and leaves unsupported payloads unchanged', () => {
		const edited = writeOleDocParagraphEdit(loadFixture(), 0, 'Line one\r\nLine two');
		expect(readOleDocParagraphs(edited)?.[0]).toBe('Line one Line two');
		const bogus = new Uint8Array([1, 2, 3]);
		expect(writeOleDocParagraphEdit(bogus, 0, 'x')).toStrictEqual(bogus);
		expect(writeOleDocParagraphEdit(loadFixture(), 99, 'x')).toStrictEqual(loadFixture());
	});
});
