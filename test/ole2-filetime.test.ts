import { describe, expect, it } from 'vitest';
import { parseOle2 } from '../src/ole2-parser-read.js';
import { buildCompoundFile, listCompoundFile, repairCompoundFile } from '../src/ole2-tree-write.js';

const EPOCH = 116444736000000000n;
const make = (createdFileTime: bigint, modifiedFileTime = createdFileTime) => buildCompoundFile([
	{ path: ['Store'], createdFileTime, modifiedFileTime },
	{ path: ['Store', 'Data'], data: new Uint8Array([1, 2, 3]), createdFileTime, modifiedFileTime },
], { rootModifiedFileTime: modifiedFileTime });

describe('lossless CFB FILETIME metadata', () => {
	it.each([0n, 1n, EPOCH - 1n, EPOCH + 12345n, 0xffffffffffffffffn])('keeps exact ticks %s through listing and repair', (ticks) => {
		const input = make(ticks);
		const listed = listCompoundFile(input);
		for (const bytes of [buildCompoundFile(listed.nodes, listed.options), repairCompoundFile(input).bytes]) {
			const parsed = parseOle2(bytes, { strict: true });
			for (const path of [['Store'], ['Store', 'Data']]) {
				const entry = parsed.findEntry(path)!;
				expect(entry.createdFileTime).toBe(ticks);
				expect(entry.modifiedFileTime).toBe(ticks);
			}
			expect(parsed.findEntry([])!.createdFileTime).toBe(0n);
			expect(parsed.findEntry([])!.modifiedFileTime).toBe(ticks);
			expect(parsed.getStreamByPath(['Store', 'Data'])).toEqual(new Uint8Array([1, 2, 3]));
		}
	});

	it('floors sub-millisecond pre-1970 dates instead of truncating toward zero', () => {
		expect(parseOle2(make(EPOCH - 1n)).findEntry('Store')!.created!.getTime()).toBe(-1);
	});

	it('honors edits to the Date view while retaining untouched precision', () => {
		const { nodes, options } = listCompoundFile(make(EPOCH + 12345n));
		nodes[0]!.created = new Date(100);
		options.rootModified = new Date(200);
		const parsed = parseOle2(buildCompoundFile(nodes, options));
		expect(parsed.findEntry('Store')!.createdFileTime).toBe(EPOCH + 1000000n);
		expect(parsed.findEntry('Store')!.modifiedFileTime).toBe(EPOCH + 12345n);
		expect(parsed.findEntry([])!.modifiedFileTime).toBe(EPOCH + 2000000n);
	});

	it('allows clearing both Date and raw views of a listed timestamp', () => {
		const { nodes, options } = listCompoundFile(make(EPOCH + 1n));
		nodes[0]!.created = undefined;
		nodes[0]!.createdFileTime = undefined;
		expect(parseOle2(buildCompoundFile(nodes, options)).findEntry('Store')!.createdFileTime).toBe(0n);
	});

	it('refuses invalid or overflowing timestamps without modifying caller data', () => {
		for (const ticks of [-1n, 0x10000000000000000n]) expect(() => make(ticks)).toThrow(/unsigned 64-bit/);
		for (const created of [new Date(NaN), new Date(-8640000000000000), new Date(8640000000000000)]) {
			const data = new Uint8Array([7]);
			expect(() => buildCompoundFile([{ path: ['Data'], data, created }])).toThrow(/timestamp/);
			expect(data).toEqual(new Uint8Array([7]));
		}
	});
});
