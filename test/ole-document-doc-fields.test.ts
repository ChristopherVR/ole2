import { describe, expect, it } from 'vitest';
import { shiftDocMainFieldCps, UnsupportedDocFieldEditError } from '../src/ole-document-doc-fields.js';
import { DocResourceLimitError } from '../src/ole-document-doc-pieces.js';

function fieldPlc(body: string, sentinel = body.length, include = true): Uint8Array {
	const markers = [...body].flatMap((char, cp) => char.charCodeAt(0) >= 0x13 && char.charCodeAt(0) <= 0x15 ? [{ cp, ch: char.charCodeAt(0) }] : []);
	const actual = include ? markers : [];
	const bytes = new Uint8Array(4 + actual.length * 6), view = new DataView(bytes.buffer);
	actual.forEach(({ cp, ch }, i) => {
		view.setInt32(i * 4, cp, true);
		bytes[(actual.length + 1) * 4 + i * 2] = ch | 0x80;
		bytes[(actual.length + 1) * 4 + i * 2 + 1] = 40 + i;
	});
	view.setInt32(actual.length * 4, sentinel, true);
	return bytes;
}

const body = 'Plain\r\u0013 Outer \u0013Inner\u0014Result\u0015 \u0014Outer result\u0015\rLast\r';
const shift = (bytes: Uint8Array, start: number, end: number, delta: number, text = body) =>
	shiftDocMainFieldCps(bytes, { fc: 0, lcb: bytes.length }, text, start, end, delta, 100)!;

describe('main-story field CP preservation', () => {
	it('shifts nested field positions while retaining reserved bits and flags', () => {
		const bytes = fieldPlc(body), n = (bytes.length - 4) / 6;
		const shifted = shift(bytes, 0, 6, 10);
		const oldView = new DataView(bytes.buffer), view = new DataView(shifted.buffer);
		for (let i = 0; i <= n; i++) expect(view.getInt32(i * 4, true)).toBe(oldView.getInt32(i * 4, true) + 10);
		expect(shifted.slice((n + 1) * 4)).toStrictEqual(bytes.slice((n + 1) * 4));
	});

	it('keeps field positions before the edit and handles shrink and undefined INT32_MAX sentinels', () => {
		const bytes = fieldPlc(body), n = (bytes.length - 4) / 6;
		const start = body.indexOf('Last');
		const shifted = shift(bytes, start, body.length, -3);
		expect(shifted.slice(0, n * 4)).toStrictEqual(bytes.slice(0, n * 4));
		const maximum = fieldPlc(body, 0x7fffffff);
		const largeSentinel = shift(maximum, 0, 6, 10);
		expect(new DataView(largeSentinel.buffer).getInt32(n * 4, true)).toBe(0x7fffffff);
		expect(new DataView(shift(bytes, 0, 6, -3).buffer).getInt32(0, true)).toBe(3);
	});

	it('refuses plain paragraphs inside a multi-paragraph field result, including omitted fields', () => {
		const text = '\u0013 QUOTE "value"\u0014\rPlain result paragraph\r\u0015\r';
		const start = text.indexOf('Plain'), end = start + 'Plain result paragraph\r'.length;
		for (const include of [true, false])
			expect(() => shift(fieldPlc(text, text.length, include), start, end, 5, text)).toThrow(UnsupportedDocFieldEditError);
		const omitted = 'Plain\r\u0013 XE "Entry"\u0015\r';
		expect(() => shift(fieldPlc(omitted, omitted.length, false), 0, 6, 5, omitted)).not.toThrow();
	});

	it('rejects truncated PLCs, invalid markers, CP order/nesting and exceeded budgets', () => {
		const bytes = fieldPlc(body);
		expect(() => shiftDocMainFieldCps(bytes, { fc: 0, lcb: bytes.length - 1 }, body, 0, 6, 10, 100)).toThrow();
		const duplicate = bytes.slice();
		new DataView(duplicate.buffer).setInt32(4, new DataView(duplicate.buffer).getInt32(0, true), true);
		expect(() => shift(duplicate, 0, 6, 1)).toThrow();
		const mismatch = bytes.slice(); mismatch[mismatch.length - 2] = 0x14;
		expect(() => shift(mismatch, 0, 6, 1)).toThrow();
		expect(() => shiftDocMainFieldCps(bytes, { fc: 0, lcb: bytes.length }, body, 0, 6, 1, 1)).toThrow(DocResourceLimitError);
		expect(() => shift(fieldPlc('Plain\r\u0014Broken\r'), 0, 6, 1, 'Plain\r\u0014Broken\r')).toThrow();
		for (const limit of [-1, NaN, Infinity, 0.5])
			expect(() => shiftDocMainFieldCps(bytes, { fc: 0, lcb: bytes.length }, body, 0, 6, 1, limit)).toThrow();
	});
});
