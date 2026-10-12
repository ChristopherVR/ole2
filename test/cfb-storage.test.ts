import { readFileSync } from 'node:fs';
import { describe, expect, it } from 'vitest';
import { parseOle2, parseCompoundFile, CfbStorage, CfbStream, UnsupportedOle2EditError,
	buildCompoundFile, listCompoundFile, type CfbEntry } from '../src/index.js';

const load = (name: string) => new Uint8Array(readFileSync(new URL(`./fixtures/${name}`, import.meta.url)));
const open = (bytes: Uint8Array) => parseOle2(bytes, { expect: 'cfb' });
const fixture = () => buildCompoundFile([
	{ path: ['Source'], clsid: new Uint8Array(16).fill(3), stateBits: 7, createdFileTime: 12345678901234567n },
	{ path: ['Source', 'Nested'] },
	{ path: ['Source', 'Nested', 'Mini'], data: new Uint8Array([1, 2, 3]) },
	{ path: ['Source', 'Large'], data: new Uint8Array(9000).fill(9) },
	{ path: ['Source', 'Empty'], data: new Uint8Array() },
	{ path: ['Destination'] },
], { rootClsid: new Uint8Array(16).fill(5), rootModifiedFileTime: 12345678901234569n });

describe('CFB storage class API', () => {
	it('navigates a typed hierarchy and detaches paths, metadata and child lists', () => {
		const document = open(fixture()), source = document.root.storage('source');
		expect(document.root).toBeInstanceOf(CfbStorage);
		expect(source.path).toEqual(['Source']);
		expect(source.children.map(entry => entry.name).sort()).toEqual(['Empty', 'Large', 'Nested']);
		const entry: CfbEntry = source.storage('Nested').children[0]!;
		if (entry.kind !== 'stream') throw new Error('Expected stream');
		expect(entry.bytes).toEqual(new Uint8Array([1, 2, 3]));
		expect(Object.isFrozen(source.path)).toBe(true);
		expect(Object.isFrozen(source.children)).toBe(true);
		const baseline = document.serialize();
		source.entry.clsid.fill(0); source.entry.path!.push('Detached');
		expect(document.serialize()).toEqual(baseline);
		expect(() => document.storage(['Source', 'Large'])).toThrow(/not-a-storage/);
		expect(() => document.stream(['Source'])).toThrow(/not-a-stream/);
	});

	it('creates empty and populated entries without implicit parent creation or aliasing', () => {
		const document = open(fixture()), root = document.root;
		const storage = root.createStorage('Added');
		const input = new Uint8Array([4, 5]);
		const stream = storage.createStream('Data', input); input.fill(0);
		expect(stream).toBeInstanceOf(CfbStream);
		expect(stream.bytes).toEqual(new Uint8Array([4, 5]));
		expect(storage.createStream('Empty').bytes).toEqual(new Uint8Array());
		expect(storage.createStorage('EmptyStorage').children).toEqual([]);
		expect(document.revision).toBe(4);
		const saved = open(document.serialize());
		expect(saved.storage(['Added', 'EmptyStorage']).children).toEqual([]);
		expect(saved.stream(['Added', 'Empty']).bytes.length).toBe(0);
		expect(() => root.createStorage('added')).toThrow(/entry-already-exists/);
		expect(() => document.createStream(['Absent', 'Child'])).toThrow(/entry-not-found/);
	});

	it('moves subtrees and updates every existing descendant handle and callback', () => {
		const document = open(fixture()), source = document.storage(['Source']);
		const nested = source.storage('Nested'), stream = nested.stream('Mini');
		const duplicate = document.stream(['source', 'nested', 'mini']);
		const oldMetadata = source.entry;
		source.rename('Renamed');
		source.move(['Destination'], 'Moved');
		expect(source.path).toEqual(['Destination', 'Moved']);
		expect(nested.path).toEqual(['Destination', 'Moved', 'Nested']);
		expect(stream.path).toEqual(['Destination', 'Moved', 'Nested', 'Mini']);
		expect(duplicate.path).toEqual(stream.path);
		expect(source.entry.clsid).toEqual(oldMetadata.clsid);
		expect(source.entry.createdFileTime).toBe(oldMetadata.createdFileTime);
		expect(source.entry.stateBits).toBe(7);
		stream.rename('mini'); // A case-only rename changes the stored spelling.
		stream.move([], 'Top');
		duplicate.bytes = new Uint8Array(5000).fill(6);
		expect(stream.bytes.length).toBe(5000);
		expect(document.root.stream('Top').bytes).toEqual(stream.bytes);
		expect(document.findEntry(['Source'])).toBeUndefined();
		expect(parseCompoundFile(document.serialize(), { strict: true }).warnings).toEqual([]);
	});

	it('invalidates deleted subtree handles permanently, including after path reuse', () => {
		const document = open(fixture()), source = document.storage(['Source']);
		const nested = source.storage('Nested'), stream = nested.stream('Mini');
		const survivor = source.stream('Large');
		nested.delete();
		expect(survivor.bytes.length).toBe(9000);
		source.createStorage('Nested').createStream('Mini', new Uint8Array([8]));
		for (const action of [() => stream.bytes, () => { stream.bytes = new Uint8Array([0]); },
			() => stream.rename('Other'), () => nested.children, () => nested.createStorage('Other'), () => nested.delete()])
			expect(action).toThrow(/stale-entry/);
		expect(source.storage('Nested').stream('Mini').bytes).toEqual(new Uint8Array([8]));
		survivor.delete();
		expect(source.children.some(entry => entry.name === 'Large')).toBe(false);
	});

	it('rejects illegal operations atomically and keeps exact no-op moves clean', () => {
		const document = open(fixture()), baseline = document.serialize();
		document.rename(['Source'], 'Source');
		document.move(['Source', 'Empty'], ['Source', 'Empty']);
		expect(document.dirty).toBe(false);
		const actions = [
			() => document.delete([]), () => document.root.rename('Root'), () => document.root.move(['Destination']),
			() => document.move(['Source'], ['Source', 'Nested', 'Cycle']),
			() => document.move(['Source'], ['Destination']),
			() => document.move(['Source'], ['Missing', 'Moved']),
			() => document.createStorage(['Source', 'Large', 'Child']),
			() => document.rename(['Source'], 'bad/name'), () => document.createStorage(['x'.repeat(32)]),
			() => document.delete(['Missing']), () => document.createStream(['Source']),
		];
		for (const action of actions) {
			expect(action).toThrow(UnsupportedOle2EditError);
			expect(document.serialize()).toEqual(baseline);
			expect(document.revision).toBe(0); expect(document.dirty).toBe(false);
		}
	});

	it.each(['ole-word-97.doc', 'sample-deck.ppt', 'xls/workbook-features.xls', 'xls/workbook-encrypted.xls'])
	('preserves every unrelated stream and directory metadata in %s', name => {
		const input = load(name), before = listCompoundFile(input), document = open(input);
		const root = document.root, added = root.createStorage('ApiFixture');
		added.createStream('Payload', new Uint8Array(7000).fill(7));
		added.createStorage('Temporary').delete();
		added.rename('ApiRenamed'); added.move([], 'ApiMoved');
		const after = listCompoundFile(document.serialize(), { strict: true });
		expect(after.options).toEqual(before.options);
		for (const node of before.nodes) expect(after.nodes.find(other => JSON.stringify(other.path) === JSON.stringify(node.path))).toEqual(node);
		const original = before.nodes.find(node => node.data !== undefined)!;
		const originalStream = document.stream(original.path);
		originalStream.move(['ApiMoved'], 'OriginalMoved');
		expect(originalStream.bytes).toEqual(original.data);
		originalStream.rename('OriginalRenamed');
		document.move(originalStream.path, original.path);
		expect(originalStream.path).toEqual(original.path);
		document.delete(['ApiMoved']);
		const restored = listCompoundFile(document.serialize(), { strict: true }).nodes;
		expect(restored.length).toBe(before.nodes.length);
		for (const node of before.nodes) expect(restored.find(other => JSON.stringify(other.path) === JSON.stringify(node.path))).toEqual(node);
	});

	it('refuses irregular input without implicitly repairing it', () => {
		const irregular = fixture(); new DataView(irregular.buffer).setUint16(0x1c, 0xffff, true);
		for (const input of [irregular]) {
			const document = open(input), baseline = document.serialize();
			for (const action of [() => document.root.createStorage('New'),
				() => document.rename(document.root.children[0]!.path, 'Changed'),
				() => document.delete(document.root.children[0]!.path)]) {
				expect(action).toThrow(/structural-/);
				expect(document.serialize()).toEqual(baseline); expect(document.dirty).toBe(false);
			}
		}
	});

	it('retains the original standalone CfbStream constructor', () => {
		let data = new Uint8Array([1]);
		const stream = new CfbStream(['Legacy'], () => data, bytes => { data = new Uint8Array(bytes); });
		stream.bytes = new Uint8Array([2]); expect(stream.bytes).toEqual(new Uint8Array([2]));
		expect(() => stream.delete()).toThrow(/unowned-stream/);
	});

	it('refuses reentrant structural and stream edits during path validation', () => {
		const document = open(fixture()), baseline = document.serialize();
		const stream = document.stream(['Source', 'Large']);
		for (const nestedEdit of [() => document.createStorage(['NestedEdit']),
			() => { stream.bytes = new Uint8Array([0]); }]) {
			const path: string[] = [];
			Object.defineProperty(path, 0, { get() { nestedEdit(); return 'Added'; } });
			expect(() => document.createStorage(path)).toThrow(/reentrant-cfb-edit/);
			expect(document.serialize()).toEqual(baseline); expect(document.revision).toBe(0);
		}
		document.createStorage(['AfterRefusal']);
		expect(document.root.storage('AfterRefusal').kind).toBe('storage');
	});
});
