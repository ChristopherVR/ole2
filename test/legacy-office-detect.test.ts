import { describe, expect, it } from 'vitest';

import { identifyLegacyOffice } from '../src/legacy-office-detect.js';
import { buildOle2 } from '../src/ole2-parser-write.js';

function cfb(...names: string[]): Uint8Array {
	const streams = new Map(names.map((name) => [name, new Uint8Array([1])]));
	return new Uint8Array(buildOle2(streams));
}

describe('identifyLegacyOffice', () => {
	it.each([
		['WordDocument', 'doc', 'partial'],
		['Workbook', 'xls', 'partial'],
		['PowerPoint Document', 'ppt', 'partial'],
		['VisioDocument', 'visio', 'unsupported'],
	])('identifies %s from its defining CFB stream', (stream, format, support) => {
		expect(identifyLegacyOffice(cfb(stream)).format).toBe(format);
		expect(identifyLegacyOffice(cfb(stream)).support).toBe(support);
	});

	it('classifies the Publisher stream pair as unsupported with a heuristic label', () => {
		const result = identifyLegacyOffice(cfb('Contents', 'Escher'));
		expect(result.format).toBe('publisher');
		expect(result.support).toBe('unsupported');
		expect(result.reason).toMatch(/likely/i);
	});

	it('does not infer a format from unrelated CFB streams', () => {
		const result = identifyLegacyOffice(cfb('RandomStream'));
		expect(result.format).toBe('unknown');
		expect(result.support).toBe('unknown');
	});

	it('reports invalid input as unknown without throwing', () => {
		expect(identifyLegacyOffice(new Uint8Array([0, 1, 2]))).toMatchObject({
			format: 'unknown',
			support: 'unknown',
			streams: [],
		});
	});
});
