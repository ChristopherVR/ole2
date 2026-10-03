import { readFileSync } from 'node:fs';
import { describe, expect, it } from 'vitest';
import { readOleDocParagraphs, tryWriteOleDocParagraphEdit, DEFAULT_OLE_DOC_PROCESSING_LIMITS } from '../src/ole-document-doc-editor.js';
import { unwrapDocBytes } from '../src/ole-document-doc-cfb.js';
import { readDocFib } from '../src/ole-document-doc-fib.js';
import { buildClxBytes, parseDocClx, DocResourceLimitError } from '../src/ole-document-doc-pieces.js';

const load = () => new Uint8Array(readFileSync(new URL('./fixtures/ole-word-97.doc', import.meta.url)));

describe('DOC processing budgets', () => {
	it('rejects small physical aliased-piece documents with huge main-story expansion before decoding', () => {
		const doc = unwrapDocBytes(load())!;
		const fib = readDocFib(doc.wordDocBytes);
		const pieceLength = 4000;
		const count = Math.floor(DEFAULT_OLE_DOC_PROCESSING_LIMITS.maxCharacters / pieceLength) + 1;
		const pieces = Array.from({ length: count }, (_, i) => ({
			cpStart: i * pieceLength, cpEnd: (i + 1) * pieceLength,
			fc: 0, compressed: true, flagsWord: 0, prm: 0,
		}));
		const clx = buildClxBytes(pieces);
		const table = new Uint8Array(doc.tableBytes.length + clx.length);
		table.set(doc.tableBytes); table.set(clx, doc.tableBytes.length);
		const word = doc.wordDocBytes.slice();
		const view = new DataView(word.buffer);
		view.setInt32(fib.ccpTextOffset, count * pieceLength, true);
		view.setUint32(fib.fibRgFcLcbOffset + 33 * 8, doc.tableBytes.length, true);
		view.setUint32(fib.fibRgFcLcbOffset + 33 * 8 + 4, clx.length, true);
		const input = doc.rewrap(word, table);
		expect(input.length).toBeLessThan(100_000);
		expect(readOleDocParagraphs(input)).toBeUndefined();
		expect(tryWriteOleDocParagraphEdit(input, 0, 'x')).toEqual({ status: 'rejected', bytes: input, reason: 'resource-limit' });
	});

	it('enforces configurable character and piece budgets for reading and editing', () => {
		const input = load();
		const doc = unwrapDocBytes(input)!;
		const fib = readDocFib(doc.wordDocBytes);
		expect(readOleDocParagraphs(input, { maxCharacters: fib.ccpText - 1 })).toBeUndefined();
		expect(readOleDocParagraphs(input, { maxCharacters: fib.ccpText })).toStrictEqual(readOleDocParagraphs(input));
		expect(readOleDocParagraphs(input, { maxPieces: 0 })).toBeUndefined();
		expect(tryWriteOleDocParagraphEdit(input, 0, 'x', { maxPieces: 0 })).toEqual({ status: 'rejected', bytes: input, reason: 'resource-limit' });
		const count = parseDocClx(doc.tableBytes, fib.clx).pieces.length;
		expect(tryWriteOleDocParagraphEdit(input, 1, 'Changed length', { maxPieces: count })).toEqual({ status: 'rejected', bytes: input, reason: 'resource-limit' });
		expect(tryWriteOleDocParagraphEdit(input, 0, 'First paragraph plain text.!', { maxCharacters: fib.ccpText })).toEqual({ status: 'rejected', bytes: input, reason: 'resource-limit' });
		expect(tryWriteOleDocParagraphEdit(input, 0, 'Other paragraph plain text.', { maxCharacters: fib.ccpText, maxPieces: 100 }).status).toBe('edited');
		expect(() => parseDocClx(doc.tableBytes, fib.clx, 0)).toThrow(DocResourceLimitError);
	});

	it('rejects invalid caller budgets without altering the input', () => {
		const input = load();
		for (const value of [-1, 0.5, NaN, Infinity]) {
			expect(readOleDocParagraphs(input, { maxCharacters: value })).toBeUndefined();
			expect(tryWriteOleDocParagraphEdit(input, 0, 'x', { maxPieces: value })).toEqual({ status: 'rejected', bytes: input, reason: 'invalid-limits' });
		}
	});
});
