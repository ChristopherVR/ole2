import { readFileSync } from 'node:fs';
import { describe, expect, it } from 'vitest';
import { parseOle2 as parseCompound } from '../src/ole2-parser-read.js';
import { parseOle2 } from '../src/ole2-document.js';
import { readPptSlideTexts } from '../src/legacy-ppt-text.js';
import { readOleDocParagraphs } from '../src/ole-document-doc-editor.js';
import { buildCompoundFile, listCompoundFile, repairCompoundFile } from '../src/ole2-tree-write.js';

const load = (name: string) => new Uint8Array(readFileSync(new URL(`./fixtures/${name}`, import.meta.url)));
const REAL = ['ole-word-97.doc', 'sample-deck.ppt', 'picture-fixture.ppt', 'xls/workbook-features.xls', 'xls/workbook-encrypted.xls'];

function streamsByPath(bytes: Uint8Array): Map<string, Uint8Array> {
	const file = parseCompound(bytes);
	return new Map(file.entries.filter((e) => e.type === 2).map((e) => [e.path!.join('/'), file.getStreamById(e.id)!]));
}

describe('buildCompoundFile', () => {
	const clsid = Uint8Array.from({ length: 16 }, (_, i) => i + 1);
	const created = new Date('2001-02-03T04:05:06.789Z');
	const modified = new Date('2024-05-06T07:08:09.123Z');

	it('writes nested storages with CLSIDs, state bits and timestamps that read back strictly', () => {
		const bytes = buildCompoundFile([
			{ path: ['ObjectPool'], clsid, stateBits: 7, created, modified },
			{ path: ['ObjectPool', '_1', '\u0001CompObj'], data: new Uint8Array([1, 2, 3]) },
			{ path: ['ObjectPool', '_1', 'Large'], data: new Uint8Array(9000).fill(5) },
			{ path: ['Empty'], data: new Uint8Array(0) },
			{ path: ['EmptyStorage'] },
		], { rootClsid: clsid, rootModified: modified });
		const file = parseCompound(bytes, { strict: true });
		expect(file.warnings).toEqual([]);
		const pool = file.findEntry('ObjectPool')!;
		expect(pool.type).toBe(1);
		expect(pool.clsid).toEqual(clsid);
		expect(pool.stateBits).toBe(7);
		expect(pool.created).toEqual(created);
		expect(pool.modified).toEqual(modified);
		expect(file.findEntry([])!.clsid).toEqual(clsid);
		expect(file.findEntry([])!.created).toBeUndefined();
		expect(file.getStreamByPath('ObjectPool/_1/\u0001CompObj')).toEqual(new Uint8Array([1, 2, 3]));
		expect(file.getStreamByPath('ObjectPool/_1/Large')).toEqual(new Uint8Array(9000).fill(5));
		expect(file.getStreamByPath('Empty')).toEqual(new Uint8Array(0));
		expect(file.findEntry('EmptyStorage')?.type).toBe(1);
	});

	it('keeps sibling trees in [MS-CFB] name order so binary-search readers find every entry', () => {
		const names = Array.from({ length: 40 }, (_, i) => `S${String.fromCharCode(65 + (i * 7) % 26)}${i}`);
		const bytes = buildCompoundFile(names.map((name, i) => ({ path: ['Store', name], data: new Uint8Array([i]) })));
		const file = parseCompound(bytes, { strict: true });
		const byId = new Map(file.entries.map((e) => [e.id, e]));
		const store = file.findEntry('Store')!;
		const key = (n: string) => [n.length, n.toUpperCase()] as const;
		const search = (target: string) => {
			let id = store.childId;
			while (id !== -1) {
				const node = byId.get(id)!;
				const [a, b] = [key(target), key(node.name)];
				const cmp = a[0] - b[0] || (a[1] < b[1] ? -1 : a[1] > b[1] ? 1 : 0);
				if (cmp === 0) return node;
				id = cmp < 0 ? node.leftSiblingId : node.rightSiblingId;
			}
			return undefined;
		};
		for (const name of names) expect(search(name)?.name).toBe(name);
	});

	it('writes DIFAT sectors for files past the header FAT capacity', () => {
		const big = new Uint8Array(110 * 128 * 512).fill(3);
		const bytes = buildCompoundFile([{ path: ['Big'], data: big }]);
		expect(new DataView(bytes.buffer).getUint32(0x48, true)).toBe(1);
		const read = parseCompound(bytes, { strict: true }).getStream('Big')!;
		// Deep toEqual on 7 MB takes seconds; compare bytes directly.
		expect(read.length).toBe(big.length);
		expect(read.every((value, i) => value === big[i])).toBe(true);
	});

	it('rejects conflicting, duplicate and illegal paths', () => {
		expect(() => buildCompoundFile([{ path: ['A'], data: new Uint8Array(1) }, { path: ['a'], data: new Uint8Array(1) }])).toThrow(/Duplicate/);
		expect(() => buildCompoundFile([{ path: ['A'], data: new Uint8Array(1) }, { path: ['A', 'B'], data: new Uint8Array(1) }])).toThrow(/is a stream/);
		expect(() => buildCompoundFile([{ path: ['A', 'B'], data: new Uint8Array(1) }, { path: ['A'], data: new Uint8Array(1) }])).toThrow(/is a storage/);
		expect(() => buildCompoundFile([{ path: ['bad/name'], data: new Uint8Array(1) }])).toThrow(/must be 1-31/);
		expect(() => buildCompoundFile([{ path: ['x'.repeat(32)], data: new Uint8Array(1) }])).toThrow(/must be 1-31/);
		expect(() => buildCompoundFile([{ path: [] }])).toThrow(/at least one/);
	});
});

describe('listCompoundFile -> edit -> buildCompoundFile', () => {
	it('adds, deletes and moves entries of a real Word file without touching other streams', () => {
		const input = load('ole-word-97.doc');
		const before = streamsByPath(input);
		const { nodes, options } = listCompoundFile(input);
		const kept = nodes.filter((node) => node.path.join('/') !== '\u0005SummaryInformation');
		kept.push({ path: ['Attachments', 'note.txt'], data: new TextEncoder().encode('hello') });
		for (const node of kept) if (node.path[0] === '\u0005DocumentSummaryInformation') node.path = ['Moved', ...node.path];
		const after = streamsByPath(buildCompoundFile(kept, options));
		expect(after.has('\u0005SummaryInformation')).toBe(false);
		expect(new TextDecoder().decode(after.get('Attachments/note.txt'))).toBe('hello');
		expect(after.get('Moved/\u0005DocumentSummaryInformation')).toEqual(before.get('\u0005DocumentSummaryInformation'));
		expect(after.get('WordDocument')).toEqual(before.get('WordDocument'));
		expect(after.get('1Table')).toEqual(before.get('1Table'));
	});
});

describe('repairCompoundFile', () => {
	it.each(REAL)('rebuilds %s canonically with identical streams and root identity', (name) => {
		const input = load(name);
		const { bytes, warnings, dropped } = repairCompoundFile(input);
		expect(warnings).toEqual([]);
		expect(dropped).toEqual([]);
		const repaired = parseCompound(bytes, { strict: true });
		expect(streamsByPath(bytes)).toEqual(streamsByPath(input));
		expect(repaired.findEntry([])!.clsid).toEqual(parseCompound(input).findEntry([])!.clsid);
	});

	it('turns quirky files into strict-clean ones that the Office models still decode', () => {
		const deck = load('sample-deck.ppt');
		const quirky = new Uint8Array(deck.length + 37);
		quirky.set(deck);
		quirky[0x1c] = 0xff;
		quirky[0x1d] = 0xff;
		const { bytes, warnings } = repairCompoundFile(quirky);
		expect(warnings.map((w) => w.code).sort()).toEqual(['byte-order', 'unaligned-length']);
		expect(() => parseCompound(bytes, { strict: true })).not.toThrow();
		expect(parseOle2(bytes).kind).toBe('ppt');
		expect(readPptSlideTexts(bytes).slides.map((s) => s.texts.map((t) => t.text))).toEqual(
			readPptSlideTexts(deck).slides.map((s) => s.texts.map((t) => t.text)),
		);
		const doc = repairCompoundFile(load('ole-word-97.doc')).bytes;
		expect(readOleDocParagraphs(doc)).toEqual(readOleDocParagraphs(load('ole-word-97.doc')));
	});

	it('refuses unreadable streams unless asked to drop them, and never invents bytes', () => {
		const input = buildCompoundFile([
			{ path: ['Good'], data: new Uint8Array(5000).fill(1) },
			{ path: ['Bad'], data: new Uint8Array(5000).fill(2) },
		], { miniStreamCutoff: 4096 });
		const file = parseCompound(input);
		const bad = file.findEntry('Bad')!;
		const dv = new DataView(input.buffer);
		const fatAt = (dv.getUint32(0x4c, true) + 1) * 512;
		dv.setUint32(fatAt + bad.startSector * 4, 0xfffffffe, true); // chain ends after one sector
		expect(() => repairCompoundFile(input)).toThrow(/Truncated stream/);
		const { bytes, dropped } = repairCompoundFile(input, { dropUnreadable: true });
		expect(dropped).toEqual([['Bad']]);
		expect(streamsByPath(bytes)).toEqual(new Map([['Good', new Uint8Array(5000).fill(1)]]));
	});
});
