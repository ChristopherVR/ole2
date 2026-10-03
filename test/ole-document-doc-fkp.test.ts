import { describe, expect, it } from 'vitest';
import { buildSingleRunPapxPage, buildSingleRunChpxPage, extractPapxBlob, extractChpxBlob, parseBteTable } from '../src/ole-document-doc-fkp.js';

const bte = { fcs: [0, 10], pns: [0] };
describe('bounded DOC formatting pages', () => {
	it('bounds repeated late-range lookups logarithmically and checks both edges', () => {
		const count = 65536, start = (count - 1) * 10;
		const page = buildSingleRunPapxPage(start, start + 10, new Uint8Array());
		let reads = 0;
		const fcs = new Proxy(Array.from({ length: count + 1 }, (_, i) => i * 10), {
			get(target, key, receiver) {
				if (typeof key === 'string' && /^\d+$/.test(key)) reads++;
				return Reflect.get(target, key, receiver);
			},
		});
		const table = { fcs, pns: new Array(count).fill(0) };
		expect(extractPapxBlob(page, table, start)).toEqual(new Uint8Array());
		expect(reads).toBeLessThan(100);
		expect(extractPapxBlob(page, table, start + 9)).toEqual(new Uint8Array());
		expect(() => extractPapxBlob(page, table, start + 10)).toThrow('not covered');
		expect(() => extractPapxBlob(page, table, -1)).toThrow('not covered');
		expect(() => extractPapxBlob(page, table, NaN)).toThrow('not covered');
	});
	it('round-trips default PAPX properties using a zero offset', () => {
		const page = buildSingleRunPapxPage(0, 10, new Uint8Array());
		expect(page[8]).toBe(0);
		expect(extractPapxBlob(page, bte, 0)).toStrictEqual(new Uint8Array());
	});
	it('rejects truncated pages, out-of-page blobs and oversized runs', () => {
		expect(() => extractPapxBlob(new Uint8Array(511), bte, 0)).toThrow();
		const page = buildSingleRunChpxPage(0, 10, new Uint8Array([1, 0]));
		page[8] = 255;
		page[510] = 20;
		expect(() => extractChpxBlob(page, bte, 0)).toThrow();
		page[511] = 255;
		expect(() => extractChpxBlob(page, bte, 0)).toThrow();
		expect(() => buildSingleRunPapxPage(0, 10, new Uint8Array(490))).toThrow();
		expect(() => parseBteTable(new Uint8Array(12), { fc: 0, lcb: 11 })).toThrow();
	});
});
