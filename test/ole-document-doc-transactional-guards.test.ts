import { readFileSync } from 'node:fs';
import { describe, expect, it } from 'vitest';
import { readOleDocParagraphs, tryWriteOleDocParagraphEdit } from '../src/ole-document-doc-editor.js';
import { unwrapDocBytes } from '../src/ole-document-doc-cfb.js';
import { readDocFib } from '../src/ole-document-doc-fib.js';
import { parseDocClx } from '../src/ole-document-doc-pieces.js';

const load = () => new Uint8Array(readFileSync(new URL('./fixtures/ole-word-97.doc', import.meta.url)));

describe('DOC transactional field and section guards', () => {
	it('refuses a CP-stable plain paragraph within a multi-paragraph omitted field', () => {
		const doc = unwrapDocBytes(load())!, fib = readDocFib(doc.wordDocBytes);
		const paragraphs = readOleDocParagraphs(load())!;
		const pieces = parseDocClx(doc.tableBytes, fib.clx).pieces;
		const word = doc.wordDocBytes.slice();
		const setCharacter = (cp: number, ch: number) => {
			const piece = pieces.find((piece) => cp >= piece.cpStart && cp < piece.cpEnd)!;
			const fc = piece.fc + (cp - piece.cpStart) * (piece.compressed ? 1 : 2);
			if (piece.compressed) word[fc] = ch;
			else new DataView(word.buffer).setUint16(fc, ch, true);
		};
		// Derived fixture: omitted field begins before paragraph 1 and ends after
		// it. No control character occurs in the paragraph being edited.
		setCharacter(0, 0x13);
		setCharacter(1, 0x14);
		setCharacter(paragraphs[0]!.length + paragraphs[1]!.length + 2, 0x15);
		const source = doc.rewrap(word, doc.tableBytes), snapshot = source.slice();
		const result = tryWriteOleDocParagraphEdit(source, 1, 'x'.repeat(paragraphs[1]!.length));
		expect(result).toEqual({ status: 'rejected', reason: 'unsupported-features', bytes: source });
		expect(result.bytes).toBe(source);
		expect(source).toStrictEqual(snapshot);
	});

	it('refuses section tables aliased by opaque FIB extensions before shifting CPs', () => {
		const doc = unwrapDocBytes(load())!, word = doc.wordDocBytes.slice();
		const fib = readDocFib(word), view = new DataView(word.buffer);
		// Pair 94 is supported opaque metadata, but must not alias mutable PlcfSed.
		view.setUint32(fib.fibRgFcLcbOffset + 94 * 8, fib.sed.fc, true);
		view.setUint32(fib.fibRgFcLcbOffset + 94 * 8 + 4, 4, true);
		const source = doc.rewrap(word, doc.tableBytes), snapshot = source.slice();
		expect(tryWriteOleDocParagraphEdit(source, 0, 'Changed length')).toEqual({
			status: 'rejected', reason: 'invalid-document', bytes: source,
		});
		expect(source).toStrictEqual(snapshot);
	});

	it('rejects negative, descending and out-of-story section CPs atomically', () => {
		const doc = unwrapDocBytes(load())!, fib = readDocFib(doc.wordDocBytes);
		const n = (fib.sed.lcb - 4) / 16;
		expect(n).toBeGreaterThan(0);
		for (const [index, cp] of [[0, -1], [n, -1], [n, fib.ccpText + 2]]) {
			const table = doc.tableBytes.slice();
			new DataView(table.buffer).setInt32(fib.sed.fc + index! * 4, cp!, true);
			const source = doc.rewrap(doc.wordDocBytes, table), snapshot = source.slice();
			const result = tryWriteOleDocParagraphEdit(source, 0, 'Changed length');
			expect(result).toEqual({ status: 'rejected', reason: 'invalid-document', bytes: source });
			expect(result.bytes).toBe(source);
			expect(source).toStrictEqual(snapshot);
		}
		const descending = doc.tableBytes.slice(), view = new DataView(descending.buffer);
		view.setInt32(fib.sed.fc, 2, true);
		view.setInt32(fib.sed.fc + n * 4, 1, true);
		const source = doc.rewrap(doc.wordDocBytes, descending);
		expect(tryWriteOleDocParagraphEdit(source, 0, 'Changed length')).toEqual({
			status: 'rejected', reason: 'invalid-document', bytes: source,
		});
	});

	it('keeps valid growth and subsequent CP-stable edits available', () => {
		const source = load(), before = readOleDocParagraphs(source)!;
		const grown = tryWriteOleDocParagraphEdit(source, 0, 'A longer plain first paragraph.');
		expect(grown.status).toBe('edited');
		const stable = tryWriteOleDocParagraphEdit(grown.bytes, 0, 'B longer plain first paragraph.');
		expect(stable.status).toBe('edited');
		expect(readOleDocParagraphs(stable.bytes)).toStrictEqual(['B longer plain first paragraph.', ...before.slice(1)]);
	});
});
