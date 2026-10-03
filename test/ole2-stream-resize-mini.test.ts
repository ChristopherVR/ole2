import { readFileSync } from 'node:fs';
import { describe, expect, it } from 'vitest';
import { buildOle2 } from '../src/ole2-parser-write.js';
import { parseOle2 } from '../src/ole2-parser-read.js';
import { readCompoundFileStream, replaceCompoundFileStream } from '../src/ole2-stream-edit.js';
import { resizeCompoundFileStream } from '../src/ole2-stream-resize.js';
import { nestedCfb } from './fixtures/nested-cfb.js';

const data = (size: number, seed = 7) => Uint8Array.from({ length: size }, (_, i) => (i * 23 + seed) & 255);
function resized(input: Uint8Array, path: string[], replacement: Uint8Array) {
	const before = input.slice(), result = resizeCompoundFileStream(input, path, replacement);
	expect(input).toEqual(before);
	expect(result.ok).toBe(true);
	if (!result.ok) throw new Error(result.reason);
	expect(readCompoundFileStream(result.bytes, path)).toEqual(replacement);
	return result.bytes;
}

function fixture(size: number) {
	return nestedCfb([
		{ path: ['Pool', 'Target'], bytes: data(size) },
		{ path: ['Target'], bytes: data(65, 3) },
		{ path: ['Pool', 'Unknown'], bytes: data(129, 5) },
		{ path: ['Opaque'], bytes: data(5500, 9) },
	]);
}

describe('preserving CFB mini-stream resizing and transitions', () => {
	it.each([0, 1, 63, 64, 65, 511, 512, 513, 4095, 4096, 4097])('resizes nested mini payload to %i bytes without changing unknown physical bytes', (size) => {
		const input = fixture(127), view = new DataView(input.buffer);
		const directory = (view.getUint32(0x30, true) + 1) * 512;
		for (let id = 0; id < 6; id++) input.fill(id + 30, directory + id * 128 + 80, directory + id * 128 + 116);
		const before = input.slice(), output = resized(input, ['Pool', 'Target'], data(size, 11));
		expect(readCompoundFileStream(output, ['Target'])).toEqual(data(65, 3));
		expect(readCompoundFileStream(output, ['Pool', 'Unknown'])).toEqual(data(129, 5));
		expect(readCompoundFileStream(output, ['Opaque'])).toEqual(data(5500, 9));
		for (let id = 0; id < 6; id++) {
			const at = directory + id * 128;
			expect(output.slice(at, at + 116)).toEqual(before.slice(at, at + 116));
			if (id !== 0 && id !== 2) expect(output.slice(at, at + 128)).toEqual(before.slice(at, at + 128));
		}
		const fatIds = Array.from({ length: view.getUint32(0x2c, true) }, (_, i) => view.getUint32(0x4c + 4 * i, true));
		const miniFatSectors = [view.getUint32(0x3c, true)];
		// All existing physical payload bytes remain exact, including the old
		// target bytes and freed root container. Only allocation metadata moves.
		for (let sector = 0; sector < input.length / 512 - 1; sector++) {
			if (fatIds.includes(sector) || miniFatSectors.includes(sector) || (sector + 1) * 512 >= directory && (sector + 1) * 512 < directory + 2 * 512) continue;
			const at = (sector + 1) * 512;
			expect(output.slice(at, at + 512)).toEqual(before.slice(at, at + 512));
		}
	});

	it.each([0, 1, 63, 64, 65, 4095, 4096, 4097])('transitions regular payload to %i bytes preserving existing mini IDs', (size) => {
		const input = fixture(5000), output = resized(input, ['Pool', 'Target'], data(size, 11));
		expect(readCompoundFileStream(output, ['Target'])).toEqual(data(65, 3));
		expect(readCompoundFileStream(output, ['Pool', 'Unknown'])).toEqual(data(129, 5));
		expect(readCompoundFileStream(output, ['Opaque'])).toEqual(data(5500, 9));
	});

	it('creates root mini stream and MiniFAT when an existing regular-only container shrinks', () => {
		const input = new Uint8Array(buildOle2(new Map([['Target', data(5000)], ['Other', data(6000, 3)]])));
		expect(new DataView(input.buffer).getUint32(0x40, true)).toBe(0);
		const output = resized(input, ['Target'], data(65));
		expect(new DataView(output.buffer).getUint32(0x40, true)).toBe(1);
		expect(readCompoundFileStream(output, ['Other'])).toEqual(data(6000, 3));
	});

	it('grows an empty mini stream in a container without root or MiniFAT allocation', () => {
		const input = new Uint8Array(buildOle2(new Map([['Target', new Uint8Array()]])));
		expect(replaceCompoundFileStream(input, ['Target'], new Uint8Array())).not.toBe(input);
		const output = resized(input, ['Target'], data(1));
		expect(new DataView(output.buffer).getUint32(0x40, true)).toBe(1);
	});

	it.each([64, 5000])('empties a %i-byte target without extending root or physical file', (size) => {
		const input = fixture(size), view = new DataView(input.buffer);
		const directory = (view.getUint32(0x30, true) + 1) * 512;
		const output = resized(input, ['Pool', 'Target'], new Uint8Array());
		expect(output.length).toBe(input.length);
		expect(output.slice(directory, directory + 128)).toEqual(input.slice(directory, directory + 128));
		expect(output.slice(0x3c, 0x44)).toEqual(input.slice(0x3c, 0x44));
		expect(readCompoundFileStream(output, ['Pool', 'Unknown'])).toEqual(data(129, 5));
	});

	it('repeatedly moves across allocation classes and through empty payloads', () => {
		let input = fixture(1);
		for (const size of [4095, 4096, 4097, 64, 0, 65, 9000, 0, 3000]) {
			input = resized(input, ['Pool', 'Target'], data(size, 11));
			expect(readCompoundFileStream(input, ['Pool', 'Unknown'])).toEqual(data(129, 5));
			expect(readCompoundFileStream(input, ['Target'])).toEqual(data(65, 3));
		}
	});

	it('extends MiniFAT beyond 128 mini entries without modifying its old physical bytes', () => {
		const streams = new Map(Array.from({ length: 128 }, (_, i) => [`S${i.toString().padStart(3, '0')}`, data(64, i)]));
		const input = new Uint8Array(buildOle2(streams)), view = new DataView(input.buffer);
		const oldMiniFat = (view.getUint32(0x3c, true) + 1) * 512;
		const output = resized(input, ['S000'], data(65, 11));
		expect(new DataView(output.buffer).getUint32(0x40, true)).toBe(2);
		expect(output.slice(oldMiniFat, oldMiniFat + 512)).toEqual(input.slice(oldMiniFat, oldMiniFat + 512));
		for (let i = 1; i < 128; i++) expect(readCompoundFileStream(output, [`S${i.toString().padStart(3, '0')}`])).toEqual(data(64, i));
	});

	it('preserves opaque partial root tail bytes and appends after a complete mini boundary', () => {
		const input = new Uint8Array(buildOle2(new Map([['Target', data(63)], ['Unknown', data(64, 9)]])));
		const view = new DataView(input.buffer), directory = (view.getUint32(0x30, true) + 1) * 512;
		const root = (view.getUint32(directory + 116, true) + 1) * 512;
		view.setUint32(directory + 120, 129, true);
		input[root + 128] = 0xa7;
		const output = resized(input, ['Target'], data(65));
		const outView = new DataView(output.buffer), newRoot = (outView.getUint32(directory + 116, true) + 1) * 512;
		expect(output[newRoot + 128]).toBe(0xa7);
		expect(outView.getUint32(directory + 128 + 116, true)).toBe(3);
		expect(readCompoundFileStream(output, ['Unknown'])).toEqual(data(64, 9));
	});

	it.each(['ole-word-97.doc', 'sample-deck.ppt', 'xls/workbook-features.xls'])('preserves real fixture streams across property-set transitions in %s', (name) => {
		let input = new Uint8Array(readFileSync(new URL(`./fixtures/${name}`, import.meta.url)));
		const parsed = parseOle2(input.buffer), entry = parsed.entries.find((entry) => entry.type === 2 && entry.name === '\u0005SummaryInformation')!;
		const original = readCompoundFileStream(input, [entry.name])!;
		const propertyView = new DataView(original.buffer, original.byteOffset, original.byteLength);
		const sections = propertyView.getUint32(24, true);
		let used = 28 + sections * 20;
		for (let section = 0; section < sections; section++) {
			const offset = propertyView.getUint32(28 + section * 20 + 16, true);
			used = Math.max(used, offset + propertyView.getUint32(offset, true));
		}
		expect(used).toBeLessThan(4096);
		expect(original.subarray(used).some((byte) => byte !== 0)).toBe(false);
		input = resized(input, [entry.name], original.slice(0, used));
		input = resized(input, [entry.name], original);
		for (const stream of parsed.entries.filter((entry) => entry.type === 2)) expect(readCompoundFileStream(input, [stream.name])).toEqual(parsed.getStream(stream.name));
	});
});

describe('mini resize malformed/capacity guards', () => {
	it('refuses mini-sector aliasing another payload', () => {
		const input = fixture(64), view = new DataView(input.buffer), directory = (view.getUint32(0x30, true) + 1) * 512;
		view.setUint32(directory + 4 * 128 + 116, view.getUint32(directory + 2 * 128 + 116, true), true);
		const result = resizeCompoundFileStream(input, ['Pool', 'Target'], data(65));
		expect(result.ok).toBe(false);
		expect(result.bytes).toBe(input);
	});

	it('refuses a root mini stream aliasing allocation metadata', () => {
		const input = fixture(64), view = new DataView(input.buffer), directory = (view.getUint32(0x30, true) + 1) * 512;
		view.setUint32(directory + 116, view.getUint32(0x3c, true), true);
		const result = resizeCompoundFileStream(input, ['Pool', 'Target'], data(65));
		expect(result.ok).toBe(false);
		expect(result.bytes).toBe(input);
	});

	it('refuses partial mini allocations in unrelated streams instead of filling absent bytes', () => {
		const input = new Uint8Array(buildOle2(new Map([['Target', data(5000)], ['Unknown', data(64)]])));
		const view = new DataView(input.buffer), directory = (view.getUint32(0x30, true) + 1) * 512;
		view.setUint32(directory + 120, 63, true);
		const result = resizeCompoundFileStream(input, ['Target'], data(65));
		expect(result.ok).toBe(false);
		expect(result.bytes).toBe(input);
	});

	it('refuses growth past header FAT capacity from a mini allocation', () => {
		const input = fixture(64), result = resizeCompoundFileStream(input, ['Pool', 'Target'], data(8 * 1024 * 1024));
		expect(result).toEqual({ ok: false, bytes: input, reason: 'allocation-limit' });
		expect(result.bytes).toBe(input);
	});

	it('does not bypass encryption/alias guards for empty payload replacement', () => {
		const input = new Uint8Array(buildOle2(new Map([['Target', new Uint8Array()], ['EncryptionInfo', data(2)]])));
		expect(replaceCompoundFileStream(input, ['Target'], new Uint8Array())).toBe(input);
		expect(resizeCompoundFileStream(input, ['Target'], data(1)).ok).toBe(false);
	});
});
