import { readFileSync } from 'node:fs';
import { describe, expect, it } from 'vitest';
import { readPptSlideTexts, editPptSlideText, PptTextError } from '../src/legacy-ppt-text.js';
import { buildPptFile } from '../src/legacy-ppt-writer.js';
import { readCompoundFileStream, replaceCompoundFileStream } from '../src/ole2-stream-edit.js';
import { buildOle2 } from '../src/ole2-parser-write.js';
import { buildPersistDirectory } from '../src/ppt/persist-directory.js';
import { buildPersistDirectoryAtom, buildUserEditAtom } from '../src/ppt/writer/persist-writer.js';
import { RT } from '../src/legacy-ppt-record-types.js';
import { readRecordOrThrow } from '../src/legacy-ppt-record-stream.js';

const load = (name = 'sample-deck.ppt') => new Uint8Array(readFileSync(new URL(`./fixtures/${name}`, import.meta.url)));
const dv = (b: Uint8Array) => new DataView(b.buffer, b.byteOffset, b.byteLength);
const simpleDeck = (text: string) => buildPptFile({ widthEmu: 9144000, heightEmu: 5143500, pictures: [], slides: [{ shapes: [{
	kind: 'shape', spt: 1, isConnector: false, anchor: { x: 0, y: 0, w: 3000000, h: 1000000 },
	text: { textType: 4, paragraphs: [{ indentLevel: 0, runs: [{ text }] }] },
}] }] });

describe('active binary PPT text and fixed-slot editing', () => {
	it('reads real fixture slide order and text without masters or notes', () => {
		const model = readPptSlideTexts(load());
		expect(model.slides).toHaveLength(7);
		expect(model.slides.map(s => s.slideId)).toEqual([256, 257, 258, 259, 260, 261, 262]);
		expect(model.slides[0].texts.map(t => t.text)).toEqual(['Project\rAtlas', 'Product Overview', 'Q2 2026']);
		expect(model.slides[6].texts.map(t => t.text)).toEqual(['Thank you', 'atlas.example.com']);
		expect(readPptSlideTexts(load('picture-fixture.ppt')).slides[0].texts[0].text).toBe('Picture fixture');
	});
	it('edits only the target bytes, preserving pictures and every surrounding byte', () => {
		const input = load(), before = input.slice();
		const result = editPptSlideText(input, { slideIndex: 0, textIndex: 1, expectedText: 'Product Overview', text: 'Product Snapshot' });
		expect(result.status).toBe('edited');
		expect(input).toEqual(before);
		expect(result.bytes.length).toBe(input.length);
		expect(result.bytes.reduce((n, b, i) => n + Number(b !== input[i]), 0)).toBe(8);
		expect(readPptSlideTexts(result.bytes).slides[0].texts[1].text).toBe('Product Snapshot');
		const original = readPptSlideTexts(input), edited = readPptSlideTexts(result.bytes);
		edited.slides[0].texts[1].text = original.slides[0].texts[1].text;
		expect(edited).toEqual(original);
	});
	it('rejects stale locations, character growth, control changes and unencodable characters explicitly', () => {
		const bytes = load();
		for (const changes of [
			{ expectedText: 'wrong' }, { text: 'Longer replacement' }, { text: 'Project Atlas' },
			{ text: 'Project\r漢字abc' }, { slideIndex: NaN }, { textIndex: 999 },
		]) {
			const result = editPptSlideText(bytes, { slideIndex: 0, textIndex: 0, expectedText: 'Project\rAtlas', text: 'Project\rOrion', ...changes });
			expect(result.status).toBe('unsupported'); expect(result.bytes).toBe(bytes);
		}
		expect(editPptSlideText(bytes, { slideIndex: 0, textIndex: 0, expectedText: 'Project\rAtlas', text: 'Project\rAtlas' })).toEqual({ status: 'unchanged', bytes });
	});
	it('reads and edits UTF-16 including surrogate pairs without evaluating content', async () => {
		const bytes = await simpleDeck('Hello 漢字😀');
		const atom = readPptSlideTexts(bytes).slides[0].texts[0];
		expect(atom.text).toBe('Hello 漢字😀'); expect(atom.encoding).toBe('utf16');
		const result = editPptSlideText(bytes, { slideIndex: 0, textIndex: 0, expectedText: atom.text, text: 'World 仮名😃' });
		expect(result.status).toBe('edited'); expect(readPptSlideTexts(result.bytes).slides[0].texts[0].text).toBe('World 仮名😃');
		expect(editPptSlideText(bytes, { slideIndex: 0, textIndex: 0, expectedText: atom.text, text: 'Hello 漢字\ud800x' }).status).toBe('unsupported');
	});
	it('follows incremental save overrides and excludes stale slide objects', async () => {
		const bytes = await simpleDeck('Hello');
		const stream = readCompoundFileStream(bytes, ['PowerPoint Document'])!;
		const current = readCompoundFileStream(bytes, ['Current User'])!;
		const oldEdit = dv(current).getUint32(16, true);
		const chain = buildPersistDirectory(dv(stream), oldEdit);
		const model = readPptSlideTexts(bytes), slide = model.slides[0];
		const oldRecord = readRecordOrThrow(dv(stream), chain.directory.get(slide.persistId)!);
		const replacement = stream.slice(oldRecord.headerOffset, oldRecord.dataOffset + oldRecord.recLen);
		const textPos = slide.texts[0].headerOffset + 8 - oldRecord.headerOffset;
		for (let i = 0; i < 5; i++) dv(replacement).setUint16(textPos + 2 * i, 'World'.charCodeAt(i), true);
		const directory = buildPersistDirectoryAtom([[slide.persistId, stream.length]]);
		const user = buildUserEditAtom({ offsetPersistDirectory: stream.length + replacement.length, docPersistIdRef: 1, maxPersistWritten: 20, lastSlideId: slide.slideId });
		dv(user).setUint32(16, oldEdit, true);
		const appended = new Uint8Array(stream.length + replacement.length + directory.length + user.length);
		appended.set(stream); appended.set(replacement, stream.length); appended.set(directory, stream.length + replacement.length);
		const newEdit = stream.length + replacement.length + directory.length;
		appended.set(user, newEdit); dv(current).setUint32(16, newEdit, true);
		const incremental = new Uint8Array(buildOle2(new Map([['PowerPoint Document', appended], ['Current User', current]])));
		expect(readPptSlideTexts(incremental).slides[0].texts.map(t => t.text)).toEqual(['World']);
	});
	it('rejects encryption and malformed framing, preserving the original edit input', async () => {
		const encrypted = await buildPptFile({ widthEmu: 9144000, heightEmu: 5143500, pictures: [], slides: [{ shapes: [] }] }, { password: 'test' });
		expect(() => readPptSlideTexts(encrypted)).toThrow(PptTextError);
		try { readPptSlideTexts(encrypted); } catch (error) { expect((error as PptTextError).code).toBe('encrypted'); }
		const source = load(), stream = readCompoundFileStream(source, ['PowerPoint Document'])!;
		const location = readPptSlideTexts(source).slides[0].texts[0].headerOffset;
		dv(stream).setUint32(location + 4, 0xffffffff, true);
		const bad = replaceCompoundFileStream(source, ['PowerPoint Document'], stream);
		expect(() => readPptSlideTexts(bad)).toThrow(PptTextError);
		const result = editPptSlideText(bad, { slideIndex: 0, textIndex: 0, expectedText: '', text: '' });
		expect(result.status).toBe('unsupported'); expect(result.bytes).toBe(bad);
	});
	it('refuses binary edits when modern OOXML content can override the slide', async () => {
		const bytes = await buildPptFile({ widthEmu: 9144000, heightEmu: 5143500, pictures: [], slides: [{ shapes: [{
			kind: 'shape', spt: 1, isConnector: false, anchor: { x: 0, y: 0, w: 3000000, h: 1000000 },
			metroBlob: new Uint8Array([80, 75, 3, 4]),
			text: { textType: 4, paragraphs: [{ indentLevel: 0, runs: [{ text: 'Hello' }] }] },
		}] }] });
		expect(readPptSlideTexts(bytes).slides[0].texts[0].text).toBe('Hello');
		const result = editPptSlideText(bytes, { slideIndex: 0, textIndex: 0, expectedText: 'Hello', text: 'World' });
		expect(result.status).toBe('unsupported'); expect(result.bytes).toBe(bytes);
	});
});
