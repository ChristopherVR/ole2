import { describe, expect, it, vi } from 'vitest';
import { parseOle2 } from '../src/ole2-parser-read.js';
import { buildOle2 } from '../src/ole2-parser-write.js';
import { readCompoundFileStream } from '../src/ole2-stream-edit.js';

describe('CFB allocation table resource bounds', () => {
	it.each(['parser', 'path reader'])('bounds malicious FAT decoding to physical sectors in the %s', (reader) => {
		const buffer = new ArrayBuffer(65 * 512);
		new Uint8Array(buffer).set(new Uint8Array(buildOle2(new Map([['Small', new Uint8Array([7])]]))));
		const view = new DataView(buffer);
		view.setUint32(0x2c, 64, true);
		for (let i = 0; i < 64; i++) view.setUint32(0x4c + i * 4, i, true);
		let decoded = 0;
		const original = DataView.prototype.getUint32;
		const spy = vi.spyOn(DataView.prototype, 'getUint32').mockImplementation(function (this: DataView, offset, littleEndian) {
			if (this.buffer === buffer && this.byteLength === 512 && this.byteOffset > 0) decoded++;
			return original.call(this, offset, littleEndian);
		});
		try {
			if (reader === 'parser') expect(() => parseOle2(buffer)).toThrow();
			else expect(readCompoundFileStream(new Uint8Array(buffer), ['Small'])).toBeUndefined();
			expect(decoded).toBeLessThanOrEqual(64);
		} finally {
			spy.mockRestore();
		}
	});

	it.each(['parser', 'path reader'])('keeps legal mini-FAT padding without decoding unused capacity in the %s', (reader) => {
		const buffer = buildOle2(new Map([['Small', new Uint8Array([7])]]));
		const view = new DataView(buffer);
		const miniFatOffset = (view.getUint32(0x3c, true) + 1) * 512;
		let decoded = 0;
		const original = DataView.prototype.getUint32;
		const spy = vi.spyOn(DataView.prototype, 'getUint32').mockImplementation(function (this: DataView, offset, littleEndian) {
			if (this.byteLength === 512 && (reader === 'parser' ? this.buffer !== buffer && this.byteOffset === 0 : this.buffer === buffer && this.byteOffset === miniFatOffset)) decoded++;
			return original.call(this, offset, littleEndian);
		});
		try {
			const stream = reader === 'parser' ? parseOle2(buffer).getStream('Small') : readCompoundFileStream(new Uint8Array(buffer), ['Small']);
			expect(stream).toEqual(new Uint8Array([7]));
			expect(decoded).toBe(1);
		} finally {
			spy.mockRestore();
		}
	});
});
