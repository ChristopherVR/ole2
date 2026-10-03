import { readFileSync } from 'node:fs';
import { describe, expect, it } from 'vitest';
import { PptDocument } from '../src/ppt-document.js';
import { UnsupportedOle2EditError } from '../src/ole2-document-base.js';
import { editPptSlideText, readPptSlideTexts } from '../src/legacy-ppt-text.js';
import { readCompoundFileStream, replaceCompoundFileStream } from '../src/ole2-stream-edit.js';
import { HEADER_TOKEN_ENCRYPTED } from '../src/legacy-ppt-record-types.js';
import { buildPptFile } from '../src/legacy-ppt-writer.js';

const load = (file = 'ppt/native-text.ppt') => new Uint8Array(readFileSync(new URL(`./fixtures/${file}`, import.meta.url)));
const simpleDeck = (text: string) => buildPptFile({ widthEmu: 9144000, heightEmu: 5143500, pictures: [], slides: [{ shapes: [{
	kind: 'shape', spt: 1, isConnector: false, anchor: { x: 0, y: 0, w: 3000000, h: 1000000 },
	text: { textType: 4, paragraphs: [{ indentLevel: 0, runs: [{ text }] }] },
}] }] });

describe('PptDocument active text model', () => {
	it('reads owned native slides and stable identity, without marking the document dirty', () => {
		const input = load(), doc = new PptDocument(input), raw = readPptSlideTexts(input);
		expect(doc.kind).toBe('ppt'); expect(doc.dirty).toBe(false);
		expect(doc.slides.map(slide => ({ slideId: slide.slideId, persistId: slide.persistId, texts: slide.texts.map(node => ({ text: node.text, encoding: node.encoding, headerOffset: node.headerOffset })) }))).toEqual(raw.slides);
		expect(doc.serialize()).toEqual(input);
	});
	it('commits ASCII and Unicode node setters through the existing native-verified editor', () => {
		const input = load(), before = input.slice(), doc = new PptDocument(input);
		const first = doc.slides[0]!.texts[0]!, second = doc.slides[1]!.texts[0]!;
		const expectedFirst = editPptSlideText(before, { slideIndex: 0, textIndex: 0, expectedText: first.text, text: 'Native title updated' });
		first.text = 'Native title updated';
		const expectedSecond = editPptSlideText(expectedFirst.bytes, { slideIndex: 1, textIndex: 0, expectedText: second.text, text: 'Unicode \u03b2 updated' });
		second.text = 'Unicode \u03b2 updated';
		expect(doc.dirty).toBe(true); expect(doc.serialize()).toEqual(expectedSecond.bytes);
		expect(input).toEqual(before); expect(first.text).toBe('Native title updated');
		expect(new PptDocument(doc.serialize()).slides[1]!.texts[0]!.text).toBe('Unicode \u03b2 updated');
		// This connects the adapter to existing native evidence; it is not an
		// independent Office execution or a complete format-fidelity assertion.
	});
	it('supports repeated setters on retained nodes and keeps exact no-ops clean', () => {
		const doc = new PptDocument(load()), node = doc.slides[0]!.texts[0]!;
		node.text = node.text; expect(doc.dirty).toBe(false);
		node.text = 'Native title updated'; node.text = 'Native title fixture';
		expect(doc.dirty).toBe(true); expect(doc.serialize()).toEqual(load());
	});
	it('owns input and serialized bytes rather than exposing mutable backing buffers', () => {
		const input = load(), original = input.slice(), doc = new PptDocument(input);
		input.fill(0); expect(doc.serialize()).toEqual(original);
		const out = doc.serialize(); out.fill(0); expect(doc.serialize()).toEqual(original);
		expect(doc.dirty).toBe(false);
	});
	it.each(['Longer replacement title', 'Native\u0000title fixture'])('refuses %s transactionally before changing dirty or model state', value => {
		const doc = new PptDocument(load()), before = doc.serialize(), node = doc.slides[0]!.texts[0]!, text = node.text;
		expect(() => { node.text = value; }).toThrow(UnsupportedOle2EditError);
		expect(doc.dirty).toBe(false); expect(node.text).toBe(text); expect(doc.serialize()).toEqual(before);
	});
	it('retains a previously successful edit when a subsequent edit fails', () => {
		const doc = new PptDocument(load()), node = doc.slides[0]!.texts[0]!;
		node.text = 'Native title updated'; const before = doc.serialize();
		expect(() => { node.text = 'Too long for the fixed text slot'; }).toThrow(UnsupportedOle2EditError);
		expect(doc.dirty).toBe(true); expect(node.text).toBe('Native title updated'); expect(doc.serialize()).toEqual(before);
	});
	it('exposes readonly identity/arrays and explicit unsupported shapes, runs and notes', () => {
		const doc = new PptDocument(load()), slide = doc.slides[0]!, node = slide.texts[0]!;
		expect(Object.isFrozen(doc.slides)).toBe(true); expect(Object.isFrozen(slide.texts)).toBe(true);
		expect(() => Object.defineProperty(node, 'slideId', { value: 999 })).toThrow();
		expect(() => Object.defineProperty(slide, 'texts', { value: [] })).toThrow();
		expect(() => slide.shapes).toThrow(UnsupportedOle2EditError);
		expect(() => slide.notes).toThrow(UnsupportedOle2EditError);
		expect(() => node.runs).toThrow(UnsupportedOle2EditError);
		expect(doc.unsupported).toContain('shape-to-text-references');
		expect(doc.capabilities.write).toEqual(['existing-text-fixed-utf16-length']);
		expect(doc.capabilities.limitations).toContain('notes');
		expect(Reflect.set(doc, 'kind', 'xls')).toBe(false);
		expect(Reflect.set(doc, 'capabilities', {})).toBe(false);
		expect(Reflect.set(doc, 'slides', [])).toBe(false);
		expect(Reflect.set(doc, 'unsupported', [])).toBe(false);
		expect(Object.isFrozen(doc.capabilities)).toBe(true);
		expect(Object.isFrozen(doc.capabilities.read)).toBe(true);
		expect(Object.isFrozen(doc.capabilities.write)).toBe(true);
		expect(Object.isFrozen(doc.capabilities.limitations)).toBe(true);
	});
	it('explicitly refuses an OOXML-mirrored source atom without changing any bytes', () => {
		const input = load('sample-deck.ppt'), doc = new PptDocument(input), node = doc.slides[0]!.texts[0]!;
		expect(node.text).toBe('Project\rAtlas');
		expect(() => { node.text = 'Project\rOrion'; }).toThrow(/OOXML mirror/);
		expect(doc.dirty).toBe(false); expect(doc.serialize()).toEqual(input);
	});
	it('rejects malformed and encrypted sources during construction', () => {
		expect(() => new PptDocument(new Uint8Array(512))).toThrow();
		const input = load(), current = readCompoundFileStream(input, ['Current User'])!;
		new DataView(current.buffer, current.byteOffset, current.byteLength).setUint32(12, HEADER_TOKEN_ENCRYPTED, true);
		const encrypted = replaceCompoundFileStream(input, ['Current User'], current);
		expect(() => new PptDocument(encrypted)).toThrow(/Encrypted/);
	});
	it('rejects changed control slots, unpaired surrogates and unencodable narrow text', async () => {
		const doc = new PptDocument(await simpleDeck('Hello\rworld')), node = doc.slides[0]!.texts[0]!, before = doc.serialize();
		expect(() => { node.text = 'Hello world'; }).toThrow(/Control/);
		expect(() => { node.text = 'Hello\r\ud800orld'; }).toThrow(/surrogate/);
		expect(doc.serialize()).toEqual(before); expect(doc.dirty).toBe(false);
		const narrow = new PptDocument(load());
		expect(() => { narrow.slides[0]!.texts[0]!.text = '\u03b2ative title fixture'; }).toThrow(/encoding/);
	});
});
