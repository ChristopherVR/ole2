import { describe, expect, it } from 'vitest';

import { inspectLegacyVisio } from '../src/legacy-visio-inspect.js';
import { buildOle2 } from '../src/ole2-parser-write.js';

const pointerOffset = 0x24;

function vsdStream(offset: number, length: number, format = 0, version = 11): Uint8Array {
	const bytes = new Uint8Array(96);
	const view = new DataView(bytes.buffer);
	bytes.set(new TextEncoder().encode('Visio (TM) Drawing\r\n'));
	view.setUint16(0x1a, version, true);
	view.setUint32(0x1c, bytes.length, true);
	view.setUint32(pointerOffset, 0x14, true);
	view.setUint32(pointerOffset + 8, offset, true);
	view.setUint32(pointerOffset + 12, length, true);
	view.setUint16(pointerOffset + (version === 5 ? 2 : 16), format, true);
	return bytes;
}

function cfb(streams: Map<string, Uint8Array>): Uint8Array {
	return new Uint8Array(buildOle2(streams));
}

describe('inspectLegacyVisio', () => {
	it('returns a validated TrailerStream pointer without claiming a parsed drawing', () => {
		const result = inspectLegacyVisio(cfb(new Map([['VisioDocument', vsdStream(60, 20, 2)]])));

		expect(result).toMatchObject({
			format: 'vsd',
			version: 11,
			trailerOffset: 60,
			trailerLength: 20,
			trailerFormat: 2,
			trailerCompressed: true,
		});
		expect(result?.reason).toMatch(/not parsed/);
	});

	it.each([
		[
			'wrong pointer type',
			vsdStream(60, 20).map((value, index, bytes) => (index === pointerOffset ? 0x13 : value)),
		],
		['out-of-bounds trailer', vsdStream(80, 40)],
		['too-short stream', new Uint8Array(30)],
	])('rejects %s', (_label, stream) => {
		expect(inspectLegacyVisio(cfb(new Map([['VisioDocument', stream]])))).toBeUndefined();
	});

	it('does not classify an arbitrary stream with the name VisioDocument', () => {
		expect(
			inspectLegacyVisio(cfb(new Map([['VisioDocument', new Uint8Array(96)]]))),
		).toBeUndefined();
	});

	it('reads the older V5 version-dependent pointer layout', () => {
		expect(
			inspectLegacyVisio(cfb(new Map([['VisioDocument', vsdStream(60, 20, 0, 5)]]))),
		).toMatchObject({ version: 5, trailerOffset: 60, trailerLength: 20 });
	});
});
