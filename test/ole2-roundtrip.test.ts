import { describe, expect, it } from 'vitest';
import { buildOle2 } from '../src/ole2-parser-write.js';
import { parseOle2 } from '../src/ole2-parser-read.js';

describe('OLE2 named stream primitives', () => {
	it.each([4096, 0])('writes empty streams with ENDOFCHAIN at cutoff %i', (cutoff) => {
		const parsed = parseOle2(buildOle2(new Map([['Empty', new Uint8Array(0)]]), undefined, cutoff));
		expect(parsed.getStream('Empty')).toEqual(new Uint8Array(0));
		expect(parsed.entries.find((entry) => entry.name === 'Empty')?.startSector).toBe(0xfffffffe);
	});
	it('round-trips regular and mini streams and the root CLSID', () => {
		const streams = new Map([
			['Small', new Uint8Array([1, 2, 3])],
			['Large', Uint8Array.from({ length: 5000 }, (_, index) => index & 0xff)],
		]);
		const clsid = Uint8Array.from({ length: 16 }, (_, index) => index + 1);
		const parsed = parseOle2(buildOle2(streams, clsid));
		expect(parsed.getStream('Small')).toEqual(streams.get('Small'));
		expect(parsed.getStream('Large')).toEqual(streams.get('Large'));
		expect(parsed.entries.find((entry) => entry.type === 5)?.clsid).toEqual(clsid);
	});
});
