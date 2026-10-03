import { readFileSync } from 'node:fs';
import { describe, expect, it } from 'vitest';
import { DocDocument } from '../src/doc-document.js';
import { Ole2DocumentError, UnsupportedOle2EditError } from '../src/ole2-document-base.js';
import { parseOle2 } from '../src/ole2-parser-read.js';
import { readCompoundFileStream } from '../src/ole2-stream-edit.js';
import { unwrapDocBytes } from '../src/ole-document-doc-cfb.js';
import { nestedCfb } from './fixtures/nested-cfb.js';

const load = (name = 'ole-word-97.doc') => new Uint8Array(readFileSync(new URL(`./fixtures/${name}`, import.meta.url)));

describe('DocDocument paragraph model', () => {
	it('reads text and commits changes through stable paragraph accessors', () => {
		const input = load(), doc = new DocDocument(input);
		const first = doc.paragraphs[0]!, third = doc.paragraphs[2]!;
		expect(doc.kind).toBe('doc');
		expect(first.text).toBe('First paragraph plain text.');
		expect(doc.dirty).toBe(false);
		first.text = first.text;
		expect(doc.dirty).toBe(false);
		third.text = 'A much longer edited paragraph. '.repeat(20);
		expect(doc.dirty).toBe(true);
		expect(third.text).toBe('A much longer edited paragraph. '.repeat(20));
		first.text = 'Other paragraph plain text.';
		expect(first.text).toBe('Other paragraph plain text.');
		expect(doc.paragraphs[0]).toBe(first);
		const reloaded = new DocDocument(doc.serialize());
		expect(reloaded.paragraphs.map((paragraph) => paragraph.text)).toStrictEqual(doc.paragraphs.map((paragraph) => paragraph.text));
		expect(reloaded.dirty).toBe(false);
		expect(input).toStrictEqual(load());
	});

	it('refuses unsafe fields and paragraph structure without committing or marking dirty', () => {
		const doc = new DocDocument(load('doc/main-field.doc'));
		const original = doc.serialize(), text = doc.paragraphs[1]!.text;
		for (const value of ['Replacement field', 'New\rparagraph', 'New\nparagraph']) {
			expect(() => { doc.paragraphs[1]!.text = value; }).toThrow(UnsupportedOle2EditError);
			expect(doc.serialize()).toStrictEqual(original);
			expect(doc.paragraphs[1]!.text).toBe(text);
			expect(doc.dirty).toBe(false);
		}
		doc.paragraphs[0]!.text = 'A longer replacement before an untouched main-story field.';
		expect(doc.paragraphs[1]!.text).toBe(text);
		expect(new DocDocument(doc.serialize()).paragraphs[1]!.text).toBe(text);
		const revision = doc.revision, edited = doc.serialize();
		expect(() => { doc.paragraphs[1]!.text = 'Rejected'; }).toThrow(UnsupportedOle2EditError);
		expect(doc.revision).toBe(revision);
		expect(doc.dirty).toBe(true);
		expect(doc.serialize()).toStrictEqual(edited);
	});

	it('freezes paragraph structure and preserves unknown nested streams through serialization', () => {
		const bytes = load(), ole = parseOle2(bytes.buffer as ArrayBuffer);
		const streams = ole.entries.filter((entry) => entry.type === 2).map((entry) => ({ path: [entry.name], bytes: ole.getStream(entry.name)! }));
		streams.push({ path: ['Custom', 'Opaque'], bytes: new Uint8Array([9, 7, 5, 3]) });
		const doc = new DocDocument(nestedCfb(streams));
		expect(Object.isFrozen(doc.paragraphs)).toBe(true);
		expect(Object.isFrozen(doc.paragraphs[0])).toBe(true);
		expect(Object.isFrozen(doc.capabilities)).toBe(true);
		expect(Object.isFrozen(doc.capabilities.write)).toBe(true);
		expect(() => { (doc as unknown as { kind: string }).kind = 'xls'; }).toThrow();
		expect(() => { (doc.capabilities.write as unknown as string[]).push('insert-paragraph'); }).toThrow();
		expect(() => { (doc.paragraphs as unknown as string[]).push('injected'); }).toThrow();
		doc.paragraphs[0]!.text = 'Longer paragraph text while preserving nested opaque data.';
		expect(readCompoundFileStream(doc.serialize(), ['Custom', 'Opaque'])).toStrictEqual(new Uint8Array([9, 7, 5, 3]));
		const serialized = doc.serialize(), word = doc.getStream('WordDocument')!;
		word.fill(0); serialized.fill(0);
		expect(new DocDocument(doc.serialize()).paragraphs[0]!.text).toBe(doc.paragraphs[0]!.text);
	});

	it('rejects valid compound containers that are not supported DOC files, including encrypted FIBs', () => {
		expect(() => new DocDocument(load('sample-deck.ppt'))).toThrow(Ole2DocumentError);
		const input = load(), doc = unwrapDocBytes(input)!, word = doc.wordDocBytes.slice();
		const view = new DataView(word.buffer);
		view.setUint16(10, view.getUint16(10, true) | 0x100, true);
		const encrypted = doc.rewrap(word, doc.tableBytes);
		expect(() => new DocDocument(encrypted)).toThrow(Ole2DocumentError);
	});
});
