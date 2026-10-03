import { readFileSync } from 'node:fs';
import { describe, expect, it } from 'vitest';
import { Ole2ParseError, parseOle2 } from '../src/ole2-parser-read.js';
import { buildOle2 } from '../src/ole2-parser-write.js';

const fixtures = new URL('./fixtures/', import.meta.url);
function loadFixture(name: string): ArrayBuffer {
	const bytes = readFileSync(new URL(name, fixtures));
	return bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength) as ArrayBuffer;
}

describe('parseOle2 against real Office compound files', () => {
	it.each(['sample-deck.ppt', 'picture-fixture.ppt', 'ole-word-97.doc', 'xls/workbook-1904.xls', 'xls/workbook-features.xls', 'xls/workbook-styles.xls', 'xls/workbook-encrypted.xls'])('reads all allocated streams in %s at their advertised lengths', (name) => {
		const ole = parseOle2(loadFixture(name));
		for (const entry of ole.entries.filter((entry) => entry.type === 2 || entry.type === 5)) {
			expect(ole.getStream(entry.name)?.length).toBe(entry.size);
		}
	});
	it('reads directory entries and regular streams from a PowerPoint file', () => {
		const ole = parseOle2(loadFixture('sample-deck.ppt'));
		expect(ole.entries.map((entry) => entry.name)).toContain('PowerPoint Document');
		const stream = ole.getStream('PowerPoint Document');
		expect(stream).toBeDefined();
		expect(stream?.length).toBeGreaterThan(10_000);
		expect(stream?.length).toBe(
			ole.entries.find((entry) => entry.name === 'PowerPoint Document')?.size,
		);
	});

	it('reads the Pictures stream when present', () => {
		const ole = parseOle2(loadFixture('picture-fixture.ppt'));
		expect(ole.getStream('Pictures')?.length).toBeGreaterThan(0);
	});

	it('returns undefined for absent streams', () => {
		expect(parseOle2(loadFixture('sample-deck.ppt')).getStream('No Such Stream')).toBeUndefined();
	});

	it('rejects non-OLE and truncated CFB headers', () => {
		expect(() => parseOle2(new Uint8Array([0x50, 0x4b, 0x03, 0x04]).buffer)).toThrow(
			Ole2ParseError,
		);
		expect(() =>
			parseOle2(new Uint8Array([0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1]).buffer),
		).toThrow();
	});
});

describe('malformed CFB allocation bounds', () => {
	it('rejects invalid sector shifts and header counts before allocation', () => {
		for (const [offset, value, width] of [[0x1e, 31, 2], [0x20, 32, 2], [0x2c, 0xffffffff, 4], [0x48, 0xffffffff, 4]]) {
			const buffer = loadFixture('sample-deck.ppt');
			const view = new DataView(buffer);
			if (width === 2) view.setUint16(offset!, value!, true);
			else view.setUint32(offset!, value!, true);
			expect(() => parseOle2(buffer)).toThrow(Ole2ParseError);
		}
	});

	it('reports magic-only truncated headers as a parser error', () => {
		expect(() => parseOle2(new Uint8Array([0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1]).buffer))
			.toThrow(Ole2ParseError);
	});

	it('refuses to truncate a real PowerPoint stream with a premature chain end', () => {
		const buffer = loadFixture('sample-deck.ppt');
		const entry = parseOle2(buffer).entries.find((entry) => entry.name === 'PowerPoint Document')!;
		const view = new DataView(buffer);
		const fatSector = view.getUint32(0x4c, true);
		view.setUint32((fatSector + 1) * 512 + entry.startSector * 4, 0xfffffffe, true);
		expect(() => parseOle2(buffer).getStream(entry.name)).toThrow('Truncated stream');
	});

	it.each([0xffffffff, 0xfffffffd])('rejects reserved FAT terminator %i', (terminator) => {
		const buffer = buildOle2(new Map([['Large', new Uint8Array(5000)]]));
		const view = new DataView(buffer);
		const fat = (view.getUint32(0x4c, true) + 1) * 512;
		view.setUint32(fat, terminator, true);
		expect(() => parseOle2(buffer).getStream('Large')).toThrow('Invalid FAT chain terminator');
	});

	it('rejects invalid directory name lengths without leaking RangeError', () => {
		const buffer = loadFixture('ole-word-97.doc');
		const view = new DataView(buffer);
		const directory = (view.getUint32(0x30, true) + 1) * 512;
		view.setUint16(directory + 64, 0xffff, true);
		expect(() => parseOle2(buffer)).toThrow(Ole2ParseError);
	});

	it('refuses zero-filled data for a mini sector outside its root allocation', () => {
		const buffer = buildOle2(new Map([['Small', new Uint8Array([7, 8, 9])]]));
		const view = new DataView(buffer);
		const directory = (view.getUint32(0x30, true) + 1) * 512;
		view.setUint32(directory + 128 + 116, 7, true);
		expect(() => parseOle2(buffer).getStream('Small')).toThrow('Mini stream sector exceeds allocation bounds');
	});

	it('refuses a mini stream whose chain is shorter than its advertised size', () => {
		const buffer = buildOle2(new Map([['Small', new Uint8Array(65)]]));
		const view = new DataView(buffer);
		const miniFat = (view.getUint32(0x3c, true) + 1) * 512;
		view.setUint32(miniFat, 0xfffffffe, true);
		expect(() => parseOle2(buffer).getStream('Small')).toThrow('Truncated mini stream');
	});

	it('reads v4 files and rejects high stream-size bits instead of truncating them', () => {
		// Minimal v4 layout: header, directory, one regular stream sector, FAT.
		const buffer = new ArrayBuffer(4 * 4096);
		const bytes = new Uint8Array(buffer), view = new DataView(buffer);
		bytes.set([0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1]);
		view.setUint16(0x1a, 4, true);
		view.setUint16(0x1c, 0xfffe, true);
		view.setUint16(0x1e, 12, true);
		view.setUint16(0x20, 6, true);
		view.setUint32(0x28, 1, true);
		view.setUint32(0x2c, 1, true);
		view.setUint32(0x38, 4096, true);
		view.setUint32(0x3c, 0xfffffffe, true);
		view.setUint32(0x44, 0xfffffffe, true);
		bytes.fill(0xff, 0x4c, 512);
		view.setUint32(0x4c, 2, true);
		for (const [id, name, type, start, size] of [[0, 'Root Entry', 5, 0xfffffffe, 0], [1, 'Large', 2, 1, 4096]] as const) {
			const at = 4096 + id * 128;
			for (let i = 0; i < name.length; i++) view.setUint16(at + 2 * i, name.charCodeAt(i), true);
			view.setUint16(at + 64, 2 * (name.length + 1), true);
			view.setUint8(at + 66, type);
			view.setUint32(at + 68, 0xffffffff, true);
			view.setUint32(at + 72, 0xffffffff, true);
			view.setUint32(at + 76, id === 0 ? 1 : 0xffffffff, true);
			view.setUint32(at + 116, start, true);
			view.setUint32(at + 120, size, true);
		}
		bytes.fill(0x7b, 8192, 12288);
		bytes.fill(0xff, 12288);
		view.setUint32(12288, 0xfffffffe, true);
		view.setUint32(12292, 0xfffffffe, true);
		view.setUint32(12296, 0xfffffffd, true);
		expect(parseOle2(buffer).getStream('Large')).toEqual(new Uint8Array(4096).fill(0x7b));
		view.setUint32(4096 + 128 + 124, 1, true);
		expect(() => parseOle2(buffer)).toThrow('Stream size exceeds supported file bounds');
	});
});
