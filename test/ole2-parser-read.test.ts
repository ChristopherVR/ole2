import { readFileSync } from 'node:fs';
import { describe, expect, it } from 'vitest';
import { Ole2ParseError, parseOle2 } from '../src/ole2-parser-read.js';

const fixtures = new URL('./fixtures/', import.meta.url);
function loadFixture(name: string): ArrayBuffer {
	const bytes = readFileSync(new URL(name, fixtures));
	return bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength) as ArrayBuffer;
}

describe('parseOle2 against real Office compound files', () => {
	it('reads directory entries and regular streams from a PowerPoint file', () => {
		const ole = parseOle2(loadFixture('sample-deck.ppt'));
		expect(ole.entries.map((entry) => entry.name)).toContain('PowerPoint Document');
		const stream = ole.getStream('PowerPoint Document');
		expect(stream).toBeDefined();
		expect(stream?.length).toBeGreaterThan(10_000);
		expect(stream?.length).toBe(ole.entries.find((entry) => entry.name === 'PowerPoint Document')?.size);
	});

	it('reads the Pictures stream when present', () => {
		const ole = parseOle2(loadFixture('picture-fixture.ppt'));
		expect(ole.getStream('Pictures')?.length).toBeGreaterThan(0);
	});

	it('returns undefined for absent streams', () => {
		expect(parseOle2(loadFixture('sample-deck.ppt')).getStream('No Such Stream')).toBeUndefined();
	});

	it('rejects non-OLE and truncated CFB headers', () => {
		expect(() => parseOle2(new Uint8Array([0x50, 0x4b, 0x03, 0x04]).buffer)).toThrow(Ole2ParseError);
		expect(() => parseOle2(new Uint8Array([0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1]).buffer)).toThrow();
	});
});
