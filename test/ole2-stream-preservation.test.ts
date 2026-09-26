import { describe, expect, it } from 'vitest';
import { buildOle2 } from '../src/ole2-parser-write.js';
import { readCompoundFileStream, replaceCompoundFileStream } from '../src/ole2-stream-edit.js';
import { nestedCfb } from './fixtures/nested-cfb.js';

describe('compound stream preservation boundaries', () => {
	it.each([17, 5000])('preserves every other byte when editing a nested %i-byte stream', (size) => {
		const source = nestedCfb([
			{ path: ['ObjectPool', 'Payload'], bytes: new Uint8Array(size).fill(0xa3) },
			{ path: ['Payload'], bytes: new Uint8Array(55).fill(0x42) },
		]);
		const original = source.slice();
		const result = replaceCompoundFileStream(
			source,
			['ObjectPool', 'Payload'],
			new Uint8Array(size).fill(0x7b),
		);
		expect(result).not.toBe(source);
		expect(source).toEqual(original);
		expect(result.length).toBe(source.length);
		expect(result.reduce((count, value, index) => count + Number(value !== source[index]), 0)).toBe(
			size,
		);
		expect(readCompoundFileStream(result, ['Payload'])).toEqual(new Uint8Array(55).fill(0x42));
		expect(readCompoundFileStream(result, ['ObjectPool', 'Payload'])).toEqual(
			new Uint8Array(size).fill(0x7b),
		);
	});
	it('refuses a regular stream aliasing the root mini-stream allocation', () => {
		const source = new Uint8Array(
			buildOle2(
				new Map([
					['Large', new Uint8Array(5000)],
					['Small', new Uint8Array(4)],
				]),
			),
		);
		const view = new DataView(source.buffer);
		const directory = (view.getUint32(0x30, true) + 1) * 512;
		const largeStart = view.getUint32(directory + 128 + 116, true);
		view.setUint32(directory + 116, largeStart, true);
		view.setUint32(directory + 120, 5000, true);
		expect(replaceCompoundFileStream(source, ['Large'], new Uint8Array(5000).fill(1))).toBe(source);
	});
	it('refuses bad magic, 64-bit sizes and allocation metadata overlapping the stream', () => {
		const good = new Uint8Array(buildOle2(new Map([['Large', new Uint8Array(5000)]])));
		const directory = (new DataView(good.buffer).getUint32(0x30, true) + 1) * 512;
		for (const mutate of [
			(bytes: Uint8Array) => {
				bytes[7] = 0;
			},
			(bytes: Uint8Array) => {
				new DataView(bytes.buffer).setUint32(directory + 128 + 124, 1, true);
			},
			(bytes: Uint8Array) => {
				const view = new DataView(bytes.buffer);
				view.setUint32(directory + 128 + 116, view.getUint32(0x4c, true), true);
			},
		]) {
			const bad = good.slice();
			mutate(bad);
			expect(replaceCompoundFileStream(bad, ['Large'], new Uint8Array(5000).fill(1))).toBe(bad);
		}
	});
});
