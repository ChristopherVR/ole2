import { describe, expect, it } from 'vitest';
import { readCompoundFileStream } from '../src/ole2-stream-edit.js';
import { resizeCompoundFileStream } from '../src/ole2-stream-resize.js';
import { v4Cfb } from './fixtures/v4-cfb.js';

const payload = (size: number, seed = 3) => Uint8Array.from({ length: size }, (_, i) => (i * 17 + seed) & 255);
const fixture = () => v4Cfb([
	{ path: ['ObjectPool', 'Target'], bytes: payload(5000) },
	{ path: ['Target'], bytes: payload(6000, 7) },
	{ path: ['ObjectPool', 'Unknown'], bytes: payload(8000, 9) },
]);
function required(input: Uint8Array, path: string[], bytes: Uint8Array) {
	const result = resizeCompoundFileStream(input, path, bytes);
	expect(result.ok).toBe(true);
	if (!result.ok) throw new Error(result.reason);
	return result.bytes;
}

describe('preserving v4 regular stream resize', () => {
	it.each([4096, 4097, 8191, 8192, 8193, 25000])('preserves header padding, hierarchy, metadata and unknown physical bytes for %i bytes', size => {
		const input = fixture(), view = new DataView(input.buffer), dir = 4096, target = dir + 2 * 128;
		// Retain opaque transaction/application data, old payload sectors, and
		// slack bytes. The synthetic pattern is not valid Office format content.
		input.fill(0x79, 512, 4096);
		for (let id = 0; id < 5; id++) input.fill(0x30 + id, dir + id * 128 + 80, dir + id * 128 + 116);
		const old = input.slice(), fatIds = Array.from({ length: view.getUint32(0x2c, true) }, (_, i) => view.getUint32(0x4c + i * 4, true));
		const output = required(input, ['ObjectPool', 'Target'], payload(size, 11));
		expect(input).toEqual(old);
		expect(readCompoundFileStream(output, ['ObjectPool', 'Target'])).toEqual(payload(size, 11));
		expect(readCompoundFileStream(output, ['Target'])).toEqual(payload(6000, 7));
		expect(readCompoundFileStream(output, ['ObjectPool', 'Unknown'])).toEqual(payload(8000, 9));
		expect(new DataView(output.buffer).getUint32(0x28, true)).toBe(1);
		// Exclude precisely the mutable allocation fields and retained FATs.
		const normalized = output.slice(0, old.length);
		normalized.set(old.subarray(0x2c, 0x30), 0x2c);
		normalized.set(old.subarray(0x4c, 512), 0x4c);
		normalized.set(old.subarray(target + 116, target + 124), target + 116);
		for (const id of fatIds) normalized.set(old.subarray((id + 1) * 4096, (id + 2) * 4096), (id + 1) * 4096);
		expect(normalized).toEqual(old);
	});
	it('resolves slots in later directory sectors without flattening or changing count', () => {
		const streams = Array.from({ length: 40 }, (_, i) => ({ path: [`S${String(i).padStart(2, '0')}`], bytes: payload(4096, i) }));
		const input = v4Cfb(streams), output = required(input, ['S39'], payload(12000, 99));
		expect(new DataView(output.buffer).getUint32(0x28, true)).toBe(2);
		for (const stream of streams) expect(readCompoundFileStream(output, stream.path)).toEqual(stream.path[0] === 'S39' ? payload(12000, 99) : stream.bytes);
	});
	it('extends 1024-entry FATs and preserves repeated growth/shrink operations', () => {
		let input = fixture();
		for (const size of [1024 * 4096, 4096, 8192]) {
			input = required(input, ['ObjectPool', 'Target'], payload(size, 55));
			expect(Buffer.from(readCompoundFileStream(input, ['ObjectPool', 'Target'])!).equals(Buffer.from(payload(size, 55)))).toBe(true);
			expect(readCompoundFileStream(input, ['ObjectPool', 'Unknown'])).toEqual(payload(8000, 9));
			const view = new DataView(input.buffer), count = view.getUint32(0x2c, true);
			expect(count).toBe(2);
			for (let i = 0; i < count; i++) {
				const id = view.getUint32(0x4c + i * 4, true), fatId = view.getUint32(0x4c + Math.floor(id / 1024) * 4, true);
				expect(view.getUint32((fatId + 1) * 4096 + id % 1024 * 4, true)).toBe(0xfffffffd);
			}
		}
	});
	it('supports same-length edits and byte-offset input views', () => {
		const bytes = fixture(), wrapper = new Uint8Array(bytes.length + 13); wrapper.set(bytes, 7);
		const input = wrapper.subarray(7, 7 + bytes.length), old = input.slice();
		const output = required(input, ['ObjectPool', 'Target'], payload(5000, 13));
		expect(readCompoundFileStream(output, ['ObjectPool', 'Target'])).toEqual(payload(5000, 13));
		expect(input).toEqual(old);
	});
});

describe('v4 malformed and unsupported inputs fail without mutation', () => {
	it.each([
		['directory count zero', (v: DataView) => v.setUint32(0x28, 0, true)],
		['directory count mismatch', (v: DataView) => v.setUint32(0x28, 2, true)],
		['wrong sector shift', (v: DataView) => v.setUint16(0x1e, 9, true)],
		['high stream size', (v: DataView) => v.setUint32(4096 + 2 * 128 + 124, 1, true)],
		['cyclic directory allocation', (v: DataView) => v.setUint32((v.getUint32(0x4c, true) + 1) * 4096, 0, true)],
		['aliased target allocation', (v: DataView) => v.setUint32(4096 + 4 * 128 + 116, v.getUint32(4096 + 2 * 128 + 116, true), true)],
		['DIFAT extension', (v: DataView) => v.setUint32(0x48, 1, true)],
	] as const)('rejects %s', (_, mutate) => {
		const input = fixture(); mutate(new DataView(input.buffer)); const old = input.slice();
		const result = resizeCompoundFileStream(input, ['ObjectPool', 'Target'], payload(9000));
		expect(result.ok).toBe(false); expect(result.bytes).toBe(input); expect(input).toEqual(old);
	});
	it('explicitly refuses transitions into mini-stream allocation', () => {
		const input = fixture();
		expect(resizeCompoundFileStream(input, ['ObjectPool', 'Target'], payload(4095))).toEqual({ ok: false, bytes: input, reason: 'unsupported-mini-transition' });
	});
	it('rejects truncated physical sectors', () => {
		const input = fixture().slice(0, -1), result = resizeCompoundFileStream(input, ['Target'], payload(9000));
		expect(result.ok).toBe(false); expect(result.bytes).toBe(input);
	});
});
