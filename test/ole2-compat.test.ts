import { readFileSync } from 'node:fs';
import { describe, expect, it } from 'vitest';
import { Ole2ParseError, parseOle2 as parseCompound } from '../src/ole2-parser-read.js';
import { parseOle2 } from '../src/ole2-document.js';
import { buildOle2 } from '../src/ole2-parser-write.js';
import { readCompoundFileStream, replaceCompoundFileStream } from '../src/ole2-stream-edit.js';
import { buildCompoundFile } from '../src/ole2-tree-write.js';

const load = (name: string) => new Uint8Array(readFileSync(new URL(`./fixtures/${name}`, import.meta.url)));
const view = (bytes: Uint8Array) => new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);

/** File offset of a directory slot, following the FAT for multi-sector directories. */
function slotOffset(bytes: Uint8Array, id: number): number {
	const dv = view(bytes);
	const fat = (sector: number) => dv.getUint32((dv.getUint32(0x4c + Math.floor(sector / 128) * 4, true) + 1) * 512 + (sector % 128) * 4, true);
	let sector = dv.getUint32(0x30, true);
	for (let i = 0; i < Math.floor(id / 4); i++) sector = fat(sector);
	return (sector + 1) * 512 + (id % 4) * 128;
}
function slotOf(bytes: Uint8Array, name: string): number {
	return slotOffset(bytes, parseCompound(bytes).entries.find((entry) => entry.name === name)!.id);
}

const FIXTURES = [
	['ole-word-97.doc', 'WordDocument', 'doc'],
	['sample-deck.ppt', 'PowerPoint Document', 'ppt'],
	['xls/workbook-features.xls', 'Workbook', 'xls'],
] as const;

type Quirk = [string, Ole2Code: string, (bytes: Uint8Array, main: string) => Uint8Array];
const QUIRKS: Quirk[] = [
	['trailing bytes after the last sector', 'unaligned-length', (b) => { const o = new Uint8Array(b.length + 100); o.set(b); o.fill(0xcc, b.length); return o; }],
	['byte-order mark 0xFFFF', 'byte-order', (b) => { const o = b.slice(); o[0x1c] = 0xff; o[0x1d] = 0xff; return o; }],
	['uninitialized v3 stream-size high bits', 'size-high-bits', (b, main) => { const o = b.slice(); view(o).setUint32(slotOf(o, main) + 124, 0xdeadbeef, true); return o; }],
	['mini-FAT header count drift', 'mini-fat-count', (b) => { const o = b.slice(); view(o).setUint32(0x40, view(o).getUint32(0x40, true) + 1, true); return o; }],
];

describe('cfb.js-parity container tolerance on real Office fixtures', () => {
	describe.each(FIXTURES)('%s', (fixture, main, kind) => {
		const original = load(fixture);
		const expected = parseCompound(original).getStream(main)!;

		it.each(QUIRKS)('accepts %s with a warning and identical stream bytes', (_label, code, mutate) => {
			const bytes = mutate(original, main);
			const file = parseCompound(bytes);
			expect(file.getStream(main)).toEqual(expected);
			expect(file.warnings.map((warning) => warning.code)).toContain(code);
			expect(readCompoundFileStream(bytes, [main])).toEqual(expected);
			expect(parseOle2(bytes).kind).toBe(kind);
		});

		it.each(QUIRKS)('rejects %s in strict mode', (_label, code, mutate) => {
			if (code === 'size-high-bits') return; // [MS-CFB] recommends ignoring these even when strict.
			expect(() => parseCompound(mutate(original, main), { strict: true })).toThrow(Ole2ParseError);
		});

		it('keeps edits strict on the same quirks', () => {
			const bytes = QUIRKS[1]![2](original, main);
			const replacement = expected.slice();
			expect(replaceCompoundFileStream(bytes, [main], replacement)).toBe(bytes);
		});

		it('reports a canonical file without warnings', () => {
			expect(parseCompound(original).warnings).toEqual([]);
			expect(() => parseCompound(original, { strict: true })).not.toThrow();
		});
	});

	it('accepts a trimmed final sector only when the trimmed bytes were padding', () => {
		const doc = load('ole-word-97.doc');
		const trimmed = doc.slice(0, doc.length - 200);
		expect(parseCompound(trimmed).getStream('WordDocument')).toEqual(parseCompound(doc).getStream('WordDocument'));
		// The deck's last sector is allocation metadata: refusing beats inventing FAT entries.
		const deck = load('sample-deck.ppt');
		expect(() => parseCompound(deck.slice(0, deck.length - 200))).toThrow(Ole2ParseError);
	});

	it('reads a trimmed final stream sector only up to the advertised size, never zero-filling', () => {
		// 5000 bytes = 10 sectors (0..9) with 392 bytes used in the last one.
		// Move sector 9 to the physical end of the file, then trim around it.
		const payload = Uint8Array.from({ length: 5000 }, (_, i) => i % 251);
		const built = new Uint8Array(buildOle2(new Map([['Large', payload]])));
		const moved = new Uint8Array(built.length + 512);
		moved.set(built);
		const newId = built.length / 512 - 1;
		moved.set(built.subarray(10 * 512, 11 * 512), (newId + 1) * 512);
		const dv = view(moved);
		const fatAt = (dv.getUint32(0x4c, true) + 1) * 512;
		dv.setUint32(fatAt + 8 * 4, newId, true);
		dv.setUint32(fatAt + 9 * 4, 0xffffffff, true);
		dv.setUint32(fatAt + newId * 4, 0xfffffffe, true);
		expect(parseCompound(moved, { strict: true }).getStream('Large')).toEqual(payload);

		const keepsUsedBytes = moved.slice(0, moved.length - 512 + 400);
		expect(parseCompound(keepsUsedBytes).getStream('Large')).toEqual(payload);
		expect(readCompoundFileStream(keepsUsedBytes, ['Large'])).toEqual(payload);

		const losesUsedBytes = moved.slice(0, moved.length - 512 + 256);
		expect(() => parseCompound(losesUsedBytes).getStream('Large')).toThrow(Ole2ParseError);
		expect(readCompoundFileStream(losesUsedBytes, ['Large'])).toBeUndefined();
	});
});

describe('tree-aware lookup', () => {
	const nested = buildCompoundFile([
		{ path: ['ObjectPool', '_1', 'WordDocument'], data: new Uint8Array([1]) },
		{ path: ['WordDocument'], data: new Uint8Array([2]) },
		{ path: ['\u0005SummaryInformation'], data: new Uint8Array([3]) },
		{ path: ['Only', 'Nested'], data: new Uint8Array([4]) },
	]);

	it('prefers the root stream over a same-named embedded stream', () => {
		expect(parseCompound(nested).getStream('WordDocument')).toEqual(new Uint8Array([2]));
		expect(parseCompound(nested).getStreamByPath('ObjectPool/_1/WordDocument')).toEqual(new Uint8Array([1]));
		expect(parseCompound(nested).getStreamByPath(['objectpool', '_1', 'worddocument'])).toEqual(new Uint8Array([1]));
	});

	it('matches names case-insensitively and without the property-set prefix', () => {
		const file = parseCompound(nested);
		expect(file.getStream('worddocument')).toEqual(new Uint8Array([2]));
		expect(file.getStreamByPath('/SummaryInformation')).toEqual(new Uint8Array([3]));
		expect(file.getStreamByPath('Root Entry/WordDocument')).toEqual(new Uint8Array([2]));
		expect(file.getStream('Nested')).toEqual(new Uint8Array([4])); // legacy flat lookup still works
		expect(file.getStreamByPath('Nested')).toBeUndefined();
		expect(file.findEntry([])?.name).toBe('Root Entry');
		expect(file.findEntry('ObjectPool/_1')?.type).toBe(1);
	});

	it('exposes directory IDs, parent links and full paths', () => {
		const file = parseCompound(nested);
		for (const entry of file.entries) {
			expect(file.entries.find((other) => other.id === entry.id)).toBe(entry);
			if (entry.parentId >= 0) expect(entry.path?.slice(0, -1)).toEqual(file.entries.find((other) => other.id === entry.parentId)!.path);
			if (entry.type === 2) expect(file.getStreamById(entry.id)).toEqual(file.getStreamByPath(entry.path!));
		}
		const embedded = file.entries.filter((entry) => entry.name === 'WordDocument');
		expect(embedded.map((entry) => entry.path)).toEqual(expect.arrayContaining([['WordDocument'], ['ObjectPool', '_1', 'WordDocument']]));
	});

	it('reads every stream of a real Word file by ID and by path', () => {
		const file = parseCompound(load('ole-word-97.doc'));
		for (const entry of file.entries.filter((e) => e.type === 2)) {
			expect(file.getStreamById(entry.id)?.length).toBe(entry.size);
			expect(file.getStreamByPath(entry.path!)?.length).toBe(entry.size);
		}
	});

	it('keeps stale unreachable slots visible but unpathed, and rejects them strictly', () => {
		const bytes = buildOle2(new Map([['A', new Uint8Array([1])], ['B', new Uint8Array([2])]]));
		const dv = new DataView(bytes);
		const dir = (dv.getUint32(0x30, true) + 1) * 512;
		// Detach the whole sibling tree under the root.
		dv.setUint32(dir + 76, 0xffffffff, true);
		const file = parseCompound(bytes);
		expect(file.warnings.filter((w) => w.code === 'unreachable-entry')).toHaveLength(2);
		expect(file.entries.filter((e) => e.type === 2).every((e) => e.path === undefined)).toBe(true);
		expect(file.getStream('A')).toEqual(new Uint8Array([1]));
		expect(() => parseCompound(bytes, { strict: true })).toThrow(Ole2ParseError);
	});

	it('skips undecodable stale slots but still rejects an undecodable reachable one', () => {
		const bytes = new Uint8Array(buildOle2(new Map([['A', new Uint8Array([1])]])));
		const dv = view(bytes);
		const dir = (dv.getUint32(0x30, true) + 1) * 512;
		// Slot 2 is free; make it a stale stream with a garbage name length.
		bytes[dir + 256 + 66] = 2;
		dv.setUint16(dir + 256 + 64, 0xffff, true);
		expect(parseCompound(bytes).warnings.map((w) => w.code)).toContain('invalid-unreachable-entry');
		dv.setUint16(dir + 128 + 64, 0xffff, true);
		expect(() => parseCompound(bytes)).toThrow('Invalid directory entry name');
	});

	it('survives a cyclic sibling tree without hanging', () => {
		const bytes = buildOle2(new Map([['A', new Uint8Array([1])], ['B', new Uint8Array([2])], ['C', new Uint8Array([3])]]));
		const dv = new DataView(bytes);
		const dir = (dv.getUint32(0x30, true) + 1) * 512;
		const rootChild = dv.getUint32(dir + 76, true);
		dv.setUint32(dir + rootChild * 128 + 68, rootChild, true);
		const file = parseCompound(bytes);
		expect(file.warnings.map((w) => w.code)).toContain('directory-link');
		expect(() => parseCompound(bytes, { strict: true })).toThrow('Cyclic directory tree');
	});
});
