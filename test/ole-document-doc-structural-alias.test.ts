import { readFileSync } from 'node:fs';
import { describe, expect, it } from 'vitest';
import { readOleDocParagraphs, tryWriteOleDocParagraphEdit } from '../src/ole-document-doc-editor.js';
import { unwrapDocBytes } from '../src/ole-document-doc-cfb.js';
import { readDocFib } from '../src/ole-document-doc-fib.js';
import { buildClxBytes } from '../src/ole-document-doc-pieces.js';
import { parseBteTable } from '../src/ole-document-doc-fkp.js';

const load = (file = 'ole-word-97.doc') => new Uint8Array(readFileSync(new URL(`./fixtures/${file}`, import.meta.url)));

/** Malformed derivative of the existing historical fixture (origin unverified): readable text aliases an
 * otherwise opaque structural region. This is not native compatibility evidence. */
function aliasedParagraph(fc: number): Uint8Array {
	const doc = unwrapDocBytes(load())!, word = doc.wordDocBytes.slice(), fib = readDocFib(word);
	word.set(new TextEncoder().encode('Alias\r'), fc);
	const clx = buildClxBytes([{ cpStart: 0, cpEnd: 6, fc, compressed: true, flagsWord: 0, prm: 0 }]);
	const table = new Uint8Array(doc.tableBytes.length + clx.length);
	table.set(doc.tableBytes); table.set(clx, doc.tableBytes.length);
	const view = new DataView(word.buffer);
	view.setInt32(fib.ccpTextOffset, 6, true);
	view.setUint32(fib.fibRgFcLcbOffset + 33 * 8, doc.tableBytes.length, true);
	view.setUint32(fib.fibRgFcLcbOffset + 33 * 8 + 4, clx.length, true);
	return doc.rewrap(word, table);
}

describe('DOC CP-stable text structural aliases', () => {
	it('refuses a readable paragraph that aliases the FIB', () => {
		const source = aliasedParagraph(0x10), snapshot = source.slice();
		expect(readOleDocParagraphs(source)).toStrictEqual(['Alias']);
		const result = tryWriteOleDocParagraphEdit(source, 0, 'Other');
		expect(result).toEqual({ status: 'rejected', reason: 'invalid-document', bytes: source });
		expect(result.bytes).toBe(source);
		expect(source).toStrictEqual(snapshot);
	});

	it('refuses readable paragraph aliases into referenced CHPX and PAPX pages', () => {
		const doc = unwrapDocBytes(load())!, fib = readDocFib(doc.wordDocBytes);
		for (const at of [fib.plcfbteChpx, fib.plcfbtePapx]) {
			const pn = parseBteTable(doc.tableBytes, at).pns[0]! & 0x3fffff;
			const source = aliasedParagraph(pn * 512 + 400), snapshot = source.slice();
			expect(readOleDocParagraphs(source)).toStrictEqual(['Alias']);
			const result = tryWriteOleDocParagraphEdit(source, 0, 'Other');
			expect(result).toEqual({ status: 'rejected', reason: 'invalid-document', bytes: source });
			expect(result.bytes).toBe(source);
			expect(source).toStrictEqual(snapshot);
		}
	});

	it('bounds formatting-page descriptors before allocating page ranges', () => {
		const doc = unwrapDocBytes(load())!, word = doc.wordDocBytes.slice(), fib = readDocFib(word);
		new DataView(word.buffer).setUint32(fib.fibRgFcLcbOffset + 12 * 8 + 4, 65_537 * 8 + 4, true);
		const source = doc.rewrap(word, doc.tableBytes);
		expect(tryWriteOleDocParagraphEdit(source, 0, 'Other paragraph plain text.')).toEqual({
			status: 'rejected', reason: 'resource-limit', bytes: source,
		});
	});

	it('continues preserving ordinary text and all table bytes in existing Word fixtures', () => {
		for (const file of ['ole-word-97.doc', 'doc/header-field.doc', 'doc/main-field.doc', 'doc/rich-runs.doc']) {
			const source = load(file), before = readOleDocParagraphs(source)!;
			const replacement = `X${before[0]!.slice(1)}`;
			const result = tryWriteOleDocParagraphEdit(source, 0, replacement);
			expect(result.status).toBe('edited');
			if (result.status !== 'edited') throw new Error(result.reason);
			expect(result.strategy).toBe('preserved-runs');
			expect(readOleDocParagraphs(result.bytes)).toStrictEqual([replacement, ...before.slice(1)]);
			expect(unwrapDocBytes(result.bytes)!.tableBytes).toStrictEqual(unwrapDocBytes(source)!.tableBytes);
		}
	});
});
