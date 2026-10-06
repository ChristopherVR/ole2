import { readFileSync } from 'node:fs';
import { describe, expect, it } from 'vitest';
import { salvagePptText } from '../src/legacy-ppt-text-salvage.js';
import { readPptSlideTexts, PptTextError } from '../src/legacy-ppt-text.js';
import { readCompoundFileStream, replaceCompoundFileStream } from '../src/ole2-stream-edit.js';
import { RT, HEADER_TOKEN_ENCRYPTED } from '../src/legacy-ppt-record-types.js';

const load = (name = 'sample-deck.ppt') => new Uint8Array(readFileSync(new URL(`./fixtures/${name}`, import.meta.url)));
const tidy = (text: string) => text.replace(/\r\n?|\v/g, '\n');
const active = (bytes: Uint8Array) => readPptSlideTexts(bytes).slides.map((slide) => slide.texts.map((atom) => tidy(atom.text)));

function editStream(bytes: Uint8Array, name: string, mutate: (stream: Uint8Array, view: DataView) => void): Uint8Array {
	const stream = readCompoundFileStream(bytes, [name])!.slice();
	mutate(stream, new DataView(stream.buffer));
	const out = replaceCompoundFileStream(bytes, [name], stream);
	expect(out).not.toBe(bytes);
	return out;
}
/** Offsets of top-level records of one type in the PowerPoint Document stream. */
function topLevel(stream: Uint8Array, type: number): number[] {
	const view = new DataView(stream.buffer, stream.byteOffset, stream.byteLength), out: number[] = [];
	for (let at = 0; at + 8 <= stream.length; at += 8 + view.getUint32(at + 4, true)) if (view.getUint16(at + 2, true) === type) out.push(at);
	return out;
}

describe('salvagePptText', () => {
	it.each(['sample-deck.ppt', 'picture-fixture.ppt', 'ppt/native-text.ppt'])('uses the exact model for healthy %s', (name) => {
		const result = salvagePptText(load(name));
		expect(result.mode).toBe('active');
		expect(result.slides).toEqual(active(load(name)));
		expect(result.diagnostics).toEqual([]);
	});

	it('recovers the live edit when Current User points at the wrong offset', () => {
		const broken = editStream(load(), 'Current User', (_s, view) => view.setUint32(16, 3, true));
		expect(() => readPptSlideTexts(broken)).toThrow();
		const result = salvagePptText(broken);
		expect(result.mode).toBe('persist-scan');
		expect(result.slides).toEqual(active(load()));
		expect(result.slides[0]).toEqual(['Project\nAtlas', 'Product Overview', 'Q2 2026']);
	});

	it('falls back to scanning slide containers when no save history survives', () => {
		const broken = editStream(load(), 'PowerPoint Document', (stream, view) => {
			for (const at of topLevel(stream, RT.UserEditAtom)) view.setUint16(at + 2, 0x6969, true);
		});
		const result = salvagePptText(broken);
		expect(result.mode).toBe('record-scan');
		expect(result.diagnostics.at(-1)).toMatch(/superseded/);
		const flat = result.slides.flat();
		for (const text of active(load()).flat()) expect(flat).toContain(text);
	});

	it('refuses encrypted documents rather than emitting ciphertext', () => {
		const encrypted = editStream(load(), 'Current User', (_s, view) => view.setUint32(12, HEADER_TOKEN_ENCRYPTED, true));
		expect(() => salvagePptText(encrypted)).toThrow(PptTextError);
	});
});
