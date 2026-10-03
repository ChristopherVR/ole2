import { readFileSync } from 'node:fs';
import { describe, expect, it } from 'vitest';
import { buildOle2 } from '../src/ole2-parser-write.js';
import { readCompoundFileStream } from '../src/ole2-stream-edit.js';
import { parseOle2 } from '../src/ole2-parser-read.js';
import { resizeCompoundFileStream } from '../src/ole2-stream-resize.js';
import { nestedCfb } from './fixtures/nested-cfb.js';

const payload = (size: number, seed = 3) => Uint8Array.from({ length: size }, (_, i) => (i * 17 + seed) & 255);
function flat(size = 5000) {
	return new Uint8Array(buildOle2(new Map([['Target', payload(size)], ['Unknown', payload(63)], ['Another', payload(5100, 7)]])));
}
function required(input: Uint8Array, path: string[], replacement: Uint8Array) {
	const result = resizeCompoundFileStream(input, path, replacement);
	expect(result.ok).toBe(true);
	if (!result.ok) throw new Error(result.reason);
	return result.bytes;
}

describe('preserving regular CFB stream resize', () => {
	it.each([4096, 4097, 5119, 5120, 5121, 25000])('changes payload size to %i while preserving all unrelated physical bytes', (size) => {
		const input = flat(), original = input.slice(), view = new DataView(input.buffer);
		const directory = (view.getUint32(0x30, true) + 1) * 512;
		// Include application-specific CLSIDs, state bits and timestamps that a
		// flatten-and-rebuild writer would lose. Target is first in sorted slots.
		input.fill(0x75, directory + 80, directory + 116);
		input.fill(0x4a, directory + 128 + 80, directory + 128 + 116);
		const before = input.slice();
		const output = required(input, ['Target'], payload(size, 9));
		expect(input).toEqual(before);
		expect(readCompoundFileStream(output, ['Target'])).toEqual(payload(size, 9));
		expect(readCompoundFileStream(output, ['Unknown'])).toEqual(payload(63));
		expect(readCompoundFileStream(output, ['Another'])).toEqual(payload(5100, 7));
		const fatIds = Array.from({ length: view.getUint32(0x2c, true) }, (_, i) => view.getUint32(0x4c + 4 * i, true));
		for (let i = 0; i < input.length; i++) {
			const sector = Math.floor(i / 512) - 1;
			const targetFields = i >= directory + 128 + 116 && i < directory + 128 + 124;
			const headerFat = (i >= 0x2c && i < 0x30) || (i >= 0x4c && i < 512);
			if (!fatIds.includes(sector) && !targetFields && !headerFat) expect(output[i]).toBe(before[i]);
		}
		expect(output.slice(directory, directory + 128)).toEqual(before.slice(directory, directory + 128));
		expect(output.slice(directory + 128, directory + 128 + 116)).toEqual(before.slice(directory + 128, directory + 128 + 116));
		expect(original.length).toBe(input.length);
	});

	it('uses full nested storage paths and preserves duplicate leaves and raw directory metadata', () => {
		const input = nestedCfb([
			{ path: ['ObjectPool', 'Target'], bytes: payload(5000) },
			{ path: ['Target'], bytes: payload(5001, 7) },
			{ path: ['ObjectPool', 'Unknown'], bytes: payload(55, 8) },
		]);
		const view = new DataView(input.buffer), directory = (view.getUint32(0x30, true) + 1) * 512;
		for (let id = 0; id < 5; id++) input.fill(0x30 + id, directory + id * 128 + 80, directory + id * 128 + 116);
		const before = input.slice();
		const output = required(input, ['ObjectPool', 'Target'], payload(12345, 9));
		expect(readCompoundFileStream(output, ['ObjectPool', 'Target'])).toEqual(payload(12345, 9));
		expect(readCompoundFileStream(output, ['Target'])).toEqual(payload(5001, 7));
		expect(readCompoundFileStream(output, ['ObjectPool', 'Unknown'])).toEqual(payload(55, 8));
		for (let id = 0; id < 5; id++) {
			const at = directory + id * 128;
			expect(output.slice(at, at + 116)).toEqual(before.slice(at, at + 116));
			if (id !== 2) expect(output.slice(at, at + 128)).toEqual(before.slice(at, at + 128));
		}
	});

	it('extends the FAT at its capacity boundary without flattening or dropping unknown streams', () => {
		const input = flat(), oldCount = new DataView(input.buffer).getUint32(0x2c, true);
		const output = required(input, ['Target'], payload(130 * 512));
		const view = new DataView(output.buffer), count = view.getUint32(0x2c, true);
		expect(count).toBeGreaterThan(oldCount);
		for (let i = 0; i < count; i++) {
			const id = view.getUint32(0x4c + 4 * i, true), table = Math.floor(id / 128), slot = id % 128;
			const fatSector = view.getUint32(0x4c + 4 * table, true);
			expect(view.getUint32((fatSector + 1) * 512 + slot * 4, true)).toBe(0xfffffffd);
		}
		expect(readCompoundFileStream(output, ['Target'])).toEqual(payload(130 * 512));
		expect(readCompoundFileStream(output, ['Unknown'])).toEqual(payload(63));
	});

	it.each([127, 128])('extends FAT safely from exactly %i physical sectors', (physical) => {
		const input = new Uint8Array(buildOle2(new Map([['Target', payload((physical - 2) * 512)]])));
		expect(input.length / 512 - 1).toBe(physical);
		const output = required(input, ['Target'], payload(4096, 9));
		expect(new DataView(output.buffer).getUint32(0x2c, true)).toBe(2);
		expect(readCompoundFileStream(output, ['Target'])).toEqual(payload(4096, 9));
	});

	it('supports repeated growth then shrink without changing unrelated payloads', () => {
		let input = flat();
		for (const size of [64000, 4096, 90000, 5000]) {
			input = required(input, ['Target'], payload(size, 9));
			expect(readCompoundFileStream(input, ['Target'])).toEqual(payload(size, 9));
			expect(readCompoundFileStream(input, ['Unknown'])).toEqual(payload(63));
		}
	});

	it('supports regular-stream empty payloads in the existing cutoff-0 layout', () => {
		let input = new Uint8Array(buildOle2(new Map([['Target', payload(3)]]), undefined, 0));
		input = required(input, ['Target'], new Uint8Array());
		expect(readCompoundFileStream(input, ['Target'])).toEqual(new Uint8Array());
		input = required(input, ['Target'], payload(55));
		expect(readCompoundFileStream(input, ['Target'])).toEqual(payload(55));
	});

	it.each([
		['ole-word-97.doc', 'WordDocument'],
		['sample-deck.ppt', 'PowerPoint Document'],
		['xls/workbook-features.xls', 'Workbook'],
	])('preserves every other stream in the real fixture %s', (fixture, target) => {
		const input = new Uint8Array(readFileSync(new URL(`./fixtures/${fixture}`, import.meta.url)));
		const stream = readCompoundFileStream(input, [target])!;
		const replacement = new Uint8Array(stream.length + 513);
		replacement.set(stream);
		const output = required(input, [target], replacement);
		expect(readCompoundFileStream(output, [target])).toEqual(replacement);
		const before = parseOle2(input.buffer.slice(input.byteOffset, input.byteOffset + input.byteLength) as ArrayBuffer);
		const after = parseOle2(output.buffer);
		for (const entry of before.entries.filter((entry) => entry.type === 2 && entry.name !== target)) {
			expect(after.getStream(entry.name)).toEqual(before.getStream(entry.name));
		}
	});
});

describe('explicit unsupported and malformed resize outcomes', () => {
	it.each([
		{ input: flat(), path: ['Missing'], replacement: payload(9000), reason: 'invalid-or-missing-stream' },
		{ input: flat(), path: ['Target'], replacement: payload(8 * 1024 * 1024), reason: 'allocation-limit' },
	])('returns unchanged input for $reason', ({ input, path, replacement, reason }) => {
		const before = input.slice(), result = resizeCompoundFileStream(input, path, replacement);
		expect(result).toEqual({ ok: false, bytes: input, reason });
		expect(result.bytes).toBe(input);
		expect(input).toEqual(before);
	});

	it.each([0x1a, 0x48, 0x1e])('refuses unsupported/malformed header field %i', (offset) => {
		const input = flat(), view = new DataView(input.buffer);
		if (offset === 0x48) view.setUint32(offset, 1, true);
		else view.setUint16(offset, offset === 0x1a ? 4 : 12, true);
		const result = resizeCompoundFileStream(input, ['Target'], payload(9000));
		expect(result.ok).toBe(false);
		expect(result.bytes).toBe(input);
	});

	it('rejects allocation aliases without changing unknown streams', () => {
		const input = flat(), view = new DataView(input.buffer);
		const directory = (view.getUint32(0x30, true) + 1) * 512;
		view.setUint32(directory + 2 * 128 + 116, view.getUint32(directory + 128 + 116, true), true);
		const result = resizeCompoundFileStream(input, ['Target'], payload(9000));
		expect(result.ok).toBe(false);
		expect(result.bytes).toBe(input);
	});

	it('rejects a target aliasing directory allocation metadata', () => {
		const input = flat(), view = new DataView(input.buffer);
		const directorySector = view.getUint32(0x30, true), directory = (directorySector + 1) * 512;
		view.setUint32(directory + 128 + 116, directorySector, true);
		const result = resizeCompoundFileStream(input, ['Target'], payload(9000));
		expect(result.ok).toBe(false);
		expect(result.bytes).toBe(input);
	});

	it.each([5000, 9000])('refuses a directory stream shared by two storage parents for %i-byte edits', (size) => {
		const input = nestedCfb([
			{ path: ['Alpha', 'Target'], bytes: payload(5000) },
			{ path: ['Beta', 'Other'], bytes: payload(5100) },
		]);
		const view = new DataView(input.buffer), directory = (view.getUint32(0x30, true) + 1) * 512;
		// Beta's child is illegally changed to Alpha's Target directory slot.
		view.setUint32(directory + 3 * 128 + 76, 2, true);
		const before = input.slice(), result = resizeCompoundFileStream(input, ['Alpha', 'Target'], payload(size, 9));
		expect(result).toEqual({ ok: false, bytes: input, reason: 'unsafe-edit' });
		expect(result.bytes).toBe(input);
		expect(input).toEqual(before);
	});

	it('rejects the actual encrypted XLS fixture', () => {
		const input = new Uint8Array(readFileSync(new URL('./fixtures/xls/workbook-encrypted.xls', import.meta.url)));
		const result = resizeCompoundFileStream(input, ['Workbook'], payload(30000));
		expect(result).toEqual({ ok: false, bytes: input, reason: 'unsafe-edit' });
	});

	it('rejects nested encrypted workbook streams as well as root workbooks', () => {
		const bytes = new Uint8Array(5000);
		bytes.set([0x2f, 0, 0, 0]);
		const input = nestedCfb([{ path: ['ObjectPool', 'Workbook'], bytes }]);
		expect(resizeCompoundFileStream(input, ['ObjectPool', 'Workbook'], payload(9000))).toEqual({ ok: false, bytes: input, reason: 'unsafe-edit' });
	});

	it.each([0x0100, 0x8000])('rejects nested encrypted/obfuscated DOC flag %i', (flag) => {
		const bytes = new Uint8Array(5000);
		new DataView(bytes.buffer).setUint16(10, flag, true);
		const input = nestedCfb([{ path: ['ObjectPool', 'WordDocument'], bytes }]);
		expect(resizeCompoundFileStream(input, ['ObjectPool', 'WordDocument'], payload(9000))).toEqual({ ok: false, bytes: input, reason: 'unsafe-edit' });
	});
});
