import { readFileSync } from 'node:fs';
import { describe, expect, it } from 'vitest';
import { readOleDocParagraphs, tryWriteOleDocParagraphEdit } from '../src/ole-document-doc-editor.js';
import { unwrapDocBytes } from '../src/ole-document-doc-cfb.js';
import { readDocFib } from '../src/ole-document-doc-fib.js';

// Synthetic Word-authored fixture, generated with macros disabled; corpus manifest records provenance/hash.
const load = () => new Uint8Array(readFileSync(new URL('./fixtures/doc/header-field.doc', import.meta.url)));

describe('DOC edits with native header/field fixture', () => {
	it('edits plain text while preserving header/field stories and all table bytes', () => {
		const input = load();
		const original = unwrapDocBytes(input)!;
		const before = readOleDocParagraphs(input)!;
		expect(before[0]).toBe('First paragraph plain text.');
		expect(before[1]).toContain('\u0013 QUOTE "Fixture field" \u0014Fixture field\u0015');
		expect(readDocFib(original.wordDocBytes).ccpOtherStories).toBeGreaterThan(0);
		const result = tryWriteOleDocParagraphEdit(input, 0, 'Other paragraph plain text.');
		expect(result.status).toBe('edited');
		if (result.status !== 'edited') throw new Error(result.reason);
		expect(result.strategy).toBe('preserved-runs');
		const updated = unwrapDocBytes(result.bytes)!;
		expect(updated.tableBytes).toStrictEqual(original.tableBytes);
		expect(readDocFib(updated.wordDocBytes)).toStrictEqual(readDocFib(original.wordDocBytes));
		expect(readOleDocParagraphs(result.bytes)).toStrictEqual(['Other paragraph plain text.', ...before.slice(1)]);
		expect(result.bytes.length).toBe(input.length);
		// Only the five characters of the target's first word differ in the entire CFB.
		const changed = result.bytes.reduce((count, value, i) => count + Number(value !== input[i]), 0);
		expect(changed).toBe(5);
	});

	it('explicitly refuses CP-shifting edits and structural field replacement', () => {
		const input = load();
		expect(tryWriteOleDocParagraphEdit(input, 0, 'Changed length')).toEqual({ status: 'rejected', bytes: input, reason: 'unsupported-features' });
		expect(tryWriteOleDocParagraphEdit(input, 1, 'x')).toEqual({ status: 'rejected', bytes: input, reason: 'unsupported-features' });
	});
});
