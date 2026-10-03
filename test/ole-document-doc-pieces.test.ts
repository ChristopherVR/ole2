import { describe, expect, it } from 'vitest';
import { buildClxBytes, decodePiecesText, parseDocClx } from '../src/ole-document-doc-pieces.js';

const piece = { cpStart: 0, cpEnd: 3, fc: 0, compressed: true, flagsWord: 0, prm: 0 };
const parse = (bytes: Uint8Array) => parseDocClx(bytes, { fc: 0, lcb: bytes.length });

describe('bounded DOC CLX parsing', () => {
	it('rejects records crossing the declared CLX even when trailing stream bytes exist', () => {
		const bytes = buildClxBytes([piece]);
		expect(() => parseDocClx(bytes, { fc: -1, lcb: bytes.length })).toThrow();
		expect(() => parseDocClx(bytes, { fc: 0, lcb: bytes.length - 1 })).toThrow();
		for (const malformed of [new Uint8Array([1, 9, 0, 2, 0]), new Uint8Array([2, 255, 255, 255, 127])])
			expect(() => parse(malformed)).toThrow();
	});

	it('rejects partial PCD records, descending CPs and dangling Prm1 references', () => {
		const bytes = buildClxBytes([piece]);
		const view = new DataView(bytes.buffer);
		view.setInt32(1, 15, true);
		expect(() => parse(bytes)).toThrow();
		view.setInt32(1, 16, true);
		view.setInt32(9, -1, true);
		expect(() => parse(bytes)).toThrow();
		expect(() => buildClxBytes([{ ...piece, prm: 1 }])).toThrow('missing Prc');
	});

	it('rejects truncated compressed and Unicode text rather than reading partial text', () => {
		expect(() => decodePiecesText(new Uint8Array([65, 66]), [piece])).toThrow();
		expect(() => decodePiecesText(new Uint8Array(5), [{ ...piece, compressed: false }])).toThrow();
		expect(decodePiecesText(new Uint8Array([65, 66, 67]), [piece])).toBe('ABC');
	});
});
