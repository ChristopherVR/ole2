import { describe, expect, it } from 'vitest';

import { inspectLegacyPublisher } from '../src/legacy-publisher-inspect.js';
import { buildOle2 } from '../src/ole2-parser-write.js';
import { nestedCfb } from './fixtures/nested-cfb.js';

function publisher(version: number, extras: string[] = []): Uint8Array {
	const contents = new Uint8Array([0xe8, 0xac, version, 0, 1, 2, 3, 4]);
	return nestedCfb([
		{ path: ['Contents'], bytes: contents },
		...extras.map((name) => ({ path: name.split('/'), bytes: new Uint8Array([1]) })),
	]);
}

describe('inspectLegacyPublisher', () => {
	it.each([
		[0x22, 'publisher-97-2000'],
		[0x2c, 'publisher-2002'],
	] as const)('validates Publisher signature and version %s', (version, expectedVersion) => {
		const extras = version === 0x2c ? ['Escher/EscherStm', 'Quill/QuillSub/CONTENTS'] : [];
		const result = inspectLegacyPublisher(publisher(version, extras));
		expect(result).toMatchObject({ format: 'pub', version: expectedVersion });
		expect(result?.reason).toMatch(/page content was not parsed/);
	});

	it('requires the additional streams for Publisher 2002', () => {
		expect(inspectLegacyPublisher(publisher(0x2c))).toBeUndefined();
		expect(inspectLegacyPublisher(publisher(0x2c, ['Escher/EscherStm']))).toBeUndefined();
	});

	it.each([
		new Uint8Array([0, 0, 0, 0]),
		new Uint8Array([0xe8, 0xac, 0x2c, 1]),
		new Uint8Array([0xe8, 0xac, 0x33, 0]),
	])('rejects invalid or unknown Contents signatures', (contents) => {
		expect(
			inspectLegacyPublisher(new Uint8Array(buildOle2(new Map([['Contents', contents]])))),
		).toBeUndefined();
	});
});
