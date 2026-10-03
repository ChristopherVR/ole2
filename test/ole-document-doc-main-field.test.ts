import { readFileSync } from 'node:fs';
import { describe, expect, it } from 'vitest';
import { readOleDocParagraphs, tryWriteOleDocParagraphEdit } from '../src/ole-document-doc-editor.js';
import { unwrapDocBytes } from '../src/ole-document-doc-cfb.js';
import { readDocFib, readFcLcbAt } from '../src/ole-document-doc-fib.js';
import { parseOle2 } from '../src/ole2-parser-read.js';
import { readCompoundFileStream } from '../src/ole2-stream-edit.js';
import { nestedCfb } from './fixtures/nested-cfb.js';

const load = () => new Uint8Array(readFileSync(new URL('./fixtures/doc/main-field.doc', import.meta.url)));

function mainFieldPlc(input: Uint8Array) {
	const doc = unwrapDocBytes(input)!;
	const fib = readDocFib(doc.wordDocBytes);
	const at = readFcLcbAt(doc.wordDocBytes, fib, 16);
	const n = (at.lcb - 4) / 6;
	const bytes = doc.tableBytes.slice(at.fc, at.fc + at.lcb);
	const view = new DataView(bytes.buffer);
	return { doc, fib, at, n, bytes, cps: Array.from({ length: n + 1 }, (_, i) => view.getInt32(i * 4, true)) };
}

describe('Word-authored main-story field preservation', () => {
	it('grows and shrinks text before a field, moving positions and retaining field bytes', () => {
		const source = load(), before = mainFieldPlc(source), paragraphs = readOleDocParagraphs(source)!;
		expect(before.fib.ccpOtherStories).toBe(0);
		expect(paragraphs[1]).toContain('\u0013 QUOTE "Fixture field" \u0014Fixture field\u0015');
		for (const text of ['A longer replacement before an untouched main-story field.', 'Short.']) {
			const result = tryWriteOleDocParagraphEdit(source, 0, text);
			expect(result.status).toBe('edited');
			const after = mainFieldPlc(result.bytes);
			const delta = text.length - paragraphs[0]!.length;
			expect(after.cps).toStrictEqual(before.cps.map((cp) => cp + delta));
			expect(after.bytes.slice((before.n + 1) * 4)).toStrictEqual(before.bytes.slice((before.n + 1) * 4));
			expect(readOleDocParagraphs(result.bytes)).toStrictEqual([text, ...paragraphs.slice(1)]);
		}
	});

	it('retains earlier field positions when editing the following paragraph and preserves nested opaque streams', () => {
		const original = load(), ole = parseOle2(original.buffer as ArrayBuffer);
		const streams = ole.entries.filter((entry) => entry.type === 2).map((entry) => ({ path: [entry.name], bytes: ole.getStream(entry.name)! }));
		streams.push({ path: ['Custom', 'Opaque'], bytes: new Uint8Array([9, 7, 5, 3]) });
		const source = nestedCfb(streams), before = mainFieldPlc(source), paragraphs = readOleDocParagraphs(source)!;
		const result = tryWriteOleDocParagraphEdit(source, 2, 'A much longer final paragraph that follows the preserved field.');
		expect(result.status).toBe('edited');
		expect(mainFieldPlc(result.bytes).cps.slice(0, -1)).toStrictEqual(before.cps.slice(0, -1));
		expect(readOleDocParagraphs(result.bytes)?.slice(0, 2)).toStrictEqual(paragraphs.slice(0, 2));
		expect(readCompoundFileStream(result.bytes, ['Custom', 'Opaque'])).toStrictEqual(new Uint8Array([9, 7, 5, 3]));
	});

	it('rejects field replacement, aliased table storage and malformed field records without changing bytes', () => {
		const source = load(), before = mainFieldPlc(source);
		expect(tryWriteOleDocParagraphEdit(source, 1, 'Replace field')).toEqual({ status: 'rejected', bytes: source, reason: 'unsupported-features' });
		const word = before.doc.wordDocBytes.slice(), view = new DataView(word.buffer);
		view.setUint32(before.fib.fibRgFcLcbOffset + 94 * 8, before.at.fc + 4, true);
		view.setUint32(before.fib.fibRgFcLcbOffset + 94 * 8 + 4, 4, true);
		const alias = before.doc.rewrap(word, before.doc.tableBytes);
		expect(tryWriteOleDocParagraphEdit(alias, 0, 'Different length')).toEqual({ status: 'rejected', bytes: alias, reason: 'invalid-document' });
		const table = before.doc.tableBytes.slice();
		new DataView(table.buffer).setInt32(before.at.fc + 4, before.cps[0]!, true);
		const duplicate = before.doc.rewrap(before.doc.wordDocBytes, table);
		expect(tryWriteOleDocParagraphEdit(duplicate, 0, 'Different length')).toEqual({ status: 'rejected', bytes: duplicate, reason: 'invalid-document' });
		expect(tryWriteOleDocParagraphEdit(source, 0, 'Different length', { maxFieldRecords: 2 })).toEqual({ status: 'rejected', bytes: source, reason: 'resource-limit' });
	});
});
