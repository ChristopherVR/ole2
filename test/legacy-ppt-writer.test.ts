import { describe, expect, it } from 'vitest';
import { parseOle2 } from '../src/ole2-parser-read.js';
import { buildPptFile } from '../src/legacy-ppt-writer.js';
import type { WDeck } from '../src/legacy-ppt-writer.js';

const emptyDeck: WDeck = {
	widthEmu: 9144000,
	heightEmu: 5143500,
	slides: [{ shapes: [] }],
	pictures: [],
};

describe('legacy binary PowerPoint writer package API', () => {
	it('writes a parseable CFB with the expected PowerPoint streams', async () => {
		const bytes = await buildPptFile(emptyDeck);
		const cfb = parseOle2(
			bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength) as ArrayBuffer,
		);
		expect(cfb.getStream('Current User')?.length).toBeGreaterThan(0);
		expect(cfb.getStream('PowerPoint Document')?.length).toBeGreaterThan(0);
	});

	it('supports password-encrypted output through the standalone API', async () => {
		const bytes = await buildPptFile(emptyDeck, { password: 'sample' });
		const cfb = parseOle2(
			bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength) as ArrayBuffer,
		);
		expect(cfb.getStream('Current User')?.length).toBeGreaterThan(0);
		expect(cfb.getStream('PowerPoint Document')?.length).toBeGreaterThan(0);
	});
});
