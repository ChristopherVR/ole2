import { readFileSync } from 'node:fs';
import { describe, expect, it } from 'vitest';
import { readOleDocParagraphs, tryWriteOleDocParagraphEdit, writeOleDocParagraphEdit } from '../src/ole-document-doc-editor.js';
import { unwrapDocBytes } from '../src/ole-document-doc-cfb.js';
import { readDocFib } from '../src/ole-document-doc-fib.js';
import { buildClxBytes, parseDocClx } from '../src/ole-document-doc-pieces.js';
import { parseOle2 } from '../src/ole2-parser-read.js';
import { nestedCfb } from './fixtures/nested-cfb.js';
import { readCompoundFileStream } from '../src/ole2-stream-edit.js';

const fixture = new URL('./fixtures/ole-word-97.doc', import.meta.url);
const paragraphs = [
	'First paragraph plain text.',
	'Second paragraph has a bold word in the middle.',
	'Third paragraph, plain again, this is the one we will edit.',
	'Fourth and final paragraph.',
];
const loadFixture = () => new Uint8Array(readFileSync(fixture));

function mutateWord(mutate: (word: Uint8Array, fib: ReturnType<typeof readDocFib>) => void): Uint8Array {
	const doc = unwrapDocBytes(loadFixture())!;
	const word = doc.wordDocBytes.slice();
	mutate(word, readDocFib(word));
	return doc.rewrap(word, doc.tableBytes);
}

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
		let updated = writeOleDocParagraphEdit(
			loadFixture(),
			1,
			'Café naïve, curly ‘quotes’. ' + 'Long text. '.repeat(100),
		);
		updated = writeOleDocParagraphEdit(updated, 3, 'Emoji fallback 😀 end.');
		expect(readOleDocParagraphs(updated)).toStrictEqual([
			paragraphs[0],
			'Café naïve, curly ‘quotes’. ' + 'Long text. '.repeat(100),
			paragraphs[2],
			'Emoji fallback 😀 end.',
		]);
	});

	it('normalizes embedded paragraph breaks and leaves unsupported payloads unchanged', () => {
		const edited = writeOleDocParagraphEdit(loadFixture(), 0, 'Line one\r\nLine two');
		expect(readOleDocParagraphs(edited)?.[0]).toBe('Line one Line two');
		const bogus = new Uint8Array([1, 2, 3]);
		expect(writeOleDocParagraphEdit(bogus, 0, 'x')).toStrictEqual(bogus);
		expect(writeOleDocParagraphEdit(loadFixture(), 99, 'x')).toStrictEqual(loadFixture());
	});

	it('keeps CLX Prc records and references on retained and replacement pieces', () => {
		const doc = unwrapDocBytes(loadFixture())!;
		const fib = readDocFib(doc.wordDocBytes);
		const pieces = parseDocClx(doc.tableBytes, fib.clx).pieces.map((piece) => ({ ...piece, prm: 1 }));
		// [MS-DOC] sprmCFBold=0x0835, operand=1, carried by Prc index zero.
		const prc = new Uint8Array([1, 3, 0, 0x35, 0x08, 1]);
		const clx = buildClxBytes(pieces, prc);
		const table = new Uint8Array(doc.tableBytes.length + clx.length);
		table.set(doc.tableBytes);
		table.set(clx, doc.tableBytes.length);
		const word = doc.wordDocBytes.slice();
		const view = new DataView(word.buffer);
		view.setUint32(fib.fibRgFcLcbOffset + 33 * 8, doc.tableBytes.length, true);
		view.setUint32(fib.fibRgFcLcbOffset + 33 * 8 + 4, clx.length, true);
		const source = doc.rewrap(word, table);
		const result = tryWriteOleDocParagraphEdit(source, 2, 'Unicode \u6f22\u5b57 \ud83d\ude00 replacement');
		expect(result.status).toBe('edited');
		const updated = unwrapDocBytes(result.bytes)!;
		const parsed = parseDocClx(updated.tableBytes, readDocFib(updated.wordDocBytes).clx);
		expect(parsed.prcBytes).toStrictEqual(prc);
		expect(parsed.pieces.every((piece) => piece.prm === 1)).toBe(true);
		expect(readOleDocParagraphs(result.bytes)).toStrictEqual([
			paragraphs[0], paragraphs[1], 'Unicode \u6f22\u5b57 \ud83d\ude00 replacement', paragraphs[3],
		]);
		const originalOle = parseOle2(source.buffer as ArrayBuffer);
		const updatedOle = parseOle2(result.bytes.buffer as ArrayBuffer);
		for (const entry of originalOle.entries) {
			if (entry.type === 2 && !['WordDocument', fib.tableStreamName].includes(entry.name))
				expect(updatedOle.getStream(entry.name)).toStrictEqual(originalOle.getStream(entry.name));
		}
	});

	it('preserves nested containers and rejects populated endnote/drawing CP tables without changing bytes', () => {
		const doc = unwrapDocBytes(loadFixture())!;
		const nested = nestedCfb([
			{ path: ['WordDocument'], bytes: doc.wordDocBytes },
			{ path: [doc.tableStreamName], bytes: doc.tableBytes },
			{ path: ['ObjectPool', 'Private'], bytes: new Uint8Array([7, 8, 9]) },
		]);
		expect(readOleDocParagraphs(nested)).toStrictEqual(paragraphs);
		const grown = tryWriteOleDocParagraphEdit(nested, 0, 'A longer nested-container paragraph replacement.');
		expect(grown.status).toBe('edited');
		expect(readOleDocParagraphs(grown.bytes)?.[0]).toBe('A longer nested-container paragraph replacement.');
		expect(readCompoundFileStream(grown.bytes, ['ObjectPool', 'Private'])).toStrictEqual(new Uint8Array([7, 8, 9]));
		const stable = tryWriteOleDocParagraphEdit(nested, 0, 'Other paragraph plain text.');
		expect(stable.status).toBe('edited');
		expect(readCompoundFileStream(stable.bytes, ['ObjectPool', 'Private'])).toStrictEqual(new Uint8Array([7, 8, 9]));
		expect(stable.bytes.length).toBe(nested.length);
		for (const index of [2, 40, 46, 48, 56, 75, 76, 89, 90]) {
			const source = mutateWord((word, fib) => new DataView(word.buffer).setUint32(fib.fibRgFcLcbOffset + index * 8 + 4, 12, true));
			expect(tryWriteOleDocParagraphEdit(source, 0, 'x')).toEqual({ status: 'rejected', bytes: source, reason: 'unsupported-features' });
		}
	});

	it('reports invalid index, structural replacement text and encrypted input explicitly', () => {
		const source = loadFixture();
		for (const index of [-1, 0.5, NaN, 99])
			expect(tryWriteOleDocParagraphEdit(source, index, 'x')).toEqual({ status: 'rejected', bytes: source, reason: 'invalid-paragraph-index' });
		expect(tryWriteOleDocParagraphEdit(source, 0, 'cell\u0007')).toEqual({ status: 'rejected', bytes: source, reason: 'invalid-text' });
		const encrypted = mutateWord((word) => {
			const view = new DataView(word.buffer);
			view.setUint16(10, view.getUint16(10, true) | 0x100, true);
		});
		expect(readOleDocParagraphs(encrypted)).toBeUndefined();
		expect(tryWriteOleDocParagraphEdit(encrypted, 0, 'x')).toEqual({ status: 'rejected', bytes: encrypted, reason: 'invalid-payload' });
	});

	it('preserves original formatting runs and CP tables for equal-length edits with other features', () => {
		const source = mutateWord((word, fib) => new DataView(word.buffer).setUint32(fib.fibRgFcLcbOffset + 17 * 8 + 4, 12, true));
		const before = unwrapDocBytes(source)!;
		const replacement = 'Other paragraph plain text.';
		expect(replacement.length).toBe(paragraphs[0]!.length);
		const result = tryWriteOleDocParagraphEdit(source, 0, replacement);
		expect(result.status).toBe('edited');
		if (result.status !== 'edited') throw new Error(result.reason);
		expect(result.strategy).toBe('preserved-runs');
		const after = unwrapDocBytes(result.bytes)!;
		expect(after.tableBytes).toStrictEqual(before.tableBytes);
		expect(result.bytes.length).toBe(source.length);
		expect(readDocFib(after.wordDocBytes)).toStrictEqual(readDocFib(before.wordDocBytes));
		expect(readOleDocParagraphs(result.bytes)).toStrictEqual([replacement, ...paragraphs.slice(1)]);
		// A replacement requiring UTF-16 cannot reuse compressed bytes or shift a field table.
		const unicode = '\u6f22'.repeat(replacement.length);
		expect(tryWriteOleDocParagraphEdit(source, 0, unicode)).toEqual({ status: 'rejected', bytes: source, reason: 'unsupported-features' });
	});

	it('recognizes effective extended FIB versions and refuses populated unknown CP tables', () => {
		const fib = readDocFib(unwrapDocBytes(loadFixture())!.wordDocBytes);
		expect(fib.nFibBase).toBe(0xc1);
		expect(fib.nFib).toBe(0x112);
		for (const index of [115, 117, 124, 125, 180]) {
			const source = mutateWord((word, info) => new DataView(word.buffer).setUint32(info.fibRgFcLcbOffset + index * 8 + 4, 12, true));
			expect(tryWriteOleDocParagraphEdit(source, 0, 'Longer replacement paragraph text.')).toEqual({ status: 'rejected', bytes: source, reason: 'unsupported-features' });
			expect(tryWriteOleDocParagraphEdit(source, 0, 'Other paragraph plain text.').status).toBe('edited');
		}
		const shortCount = mutateWord((word, info) => new DataView(word.buffer).setUint16(info.fibRgFcLcbOffset - 2, 34, true));
		expect(readOleDocParagraphs(shortCount)).toBeUndefined();
		const badWord = unwrapDocBytes(loadFixture())!.wordDocBytes.slice();
		new DataView(badWord.buffer).setUint16(0x20, 15, true);
		expect(() => readDocFib(badWord)).toThrow();
	});
});
