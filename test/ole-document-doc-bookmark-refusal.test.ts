import { readFileSync } from 'node:fs';
import { describe, expect, it } from 'vitest';
import { DocDocument } from '../src/doc-document.js';
import { UnsupportedOle2EditError } from '../src/ole2-document-base.js';
import { unwrapDocBytes } from '../src/ole-document-doc-cfb.js';
import { readDocFib } from '../src/ole-document-doc-fib.js';
import { readOleDocParagraphs, tryWriteOleDocParagraphEdit } from '../src/ole-document-doc-editor.js';

/** Standard bookmark tables authored in-test from the primary MS-DOC layouts.
 * The base fixture is owned synthetic Word content. These derived bytes have
 * not been opened in native Word and are evidence for refusal/preservation only. */
function withBookmark(): Uint8Array {
	const source = new Uint8Array(readFileSync(new URL('./fixtures/doc/rich-runs.doc', import.meta.url)));
	const doc = unwrapDocBytes(source)!, word = doc.wordDocBytes.slice(), fib = readDocFib(word);
	const paragraphs = readOleDocParagraphs(source)!;
	const start = paragraphs[0]!.length + 1, end = start + 5;
	const name = 'OwnedBookmark';
	const names = new Uint8Array(8 + name.length * 2), nameView = new DataView(names.buffer);
	nameView.setUint16(0, 0xffff, true); nameView.setUint16(2, 1, true);
	nameView.setUint16(4, 0, true); nameView.setUint16(6, name.length, true);
	for (let i = 0; i < name.length; i++) nameView.setUint16(8 + i * 2, name.charCodeAt(i), true);
	const bkf = new Uint8Array(12), bkfView = new DataView(bkf.buffer);
	bkfView.setUint32(0, start, true); bkfView.setUint32(4, fib.ccpText + 1, true);
	bkfView.setUint16(8, 0, true); // unique ibkl: bookmark ends at BKL index zero
	bkfView.setUint16(10, 0x4000, true); // BKC.fNative; fCol/fPub clear
	const bkl = new Uint8Array(8), bklView = new DataView(bkl.buffer);
	bklView.setUint32(0, end, true); bklView.setUint32(4, fib.ccpText + 1, true);
	const table = new Uint8Array(doc.tableBytes.length + names.length + bkf.length + bkl.length);
	table.set(doc.tableBytes);
	const wordView = new DataView(word.buffer);
	let offset = doc.tableBytes.length;
	for (const [index, bytes] of [[21, names], [22, bkf], [23, bkl]] as const) {
		table.set(bytes, offset);
		wordView.setUint32(fib.fibRgFcLcbOffset + index * 8, offset, true);
		wordView.setUint32(fib.fibRgFcLcbOffset + index * 8 + 4, bytes.length, true);
		offset += bytes.length;
	}
	return doc.rewrap(word, table);
}

describe('standard DOC bookmark write boundary', () => {
	it('refuses growth before and after the bookmark even when no endpoint is replaced', () => {
		const source = withBookmark(), snapshot = source.slice();
		for (const index of [0, 2]) {
			const result = tryWriteOleDocParagraphEdit(source, index, 'Longer text outside the untouched bookmark range.');
			expect(result).toEqual({ status: 'rejected', reason: 'unsupported-features', bytes: source });
			expect(result.bytes).toBe(source);
			expect(source).toStrictEqual(snapshot);
		}
	});

	it('preserves model bytes, revision and paragraph handles on clean and dirty refusals', () => {
		const source = withBookmark(), model = new DocDocument(source);
		const first = model.paragraphs[0]!, bookmarked = model.paragraphs[1]!;
		const original = first.text, bookmarkText = bookmarked.text;
		const refuse = () => {
			const bytes = model.serialize(), revision = model.revision, dirty = model.dirty, text = first.text;
			expect(() => { first.text = `${text} plus growth`; }).toThrow(UnsupportedOle2EditError);
			expect(model.serialize()).toStrictEqual(bytes);
			expect(model.revision).toBe(revision); expect(model.dirty).toBe(dirty);
			expect(model.paragraphs[0]).toBe(first); expect(model.paragraphs[1]).toBe(bookmarked);
			expect(first.text).toBe(text); expect(bookmarked.text).toBe(bookmarkText);
		};
		refuse();
		first.text = `X${original.slice(1)}`;
		expect(model.dirty).toBe(true); expect(model.revision).toBe(1);
		expect(unwrapDocBytes(model.serialize())!.tableBytes).toStrictEqual(unwrapDocBytes(source)!.tableBytes);
		refuse();
	});
});
