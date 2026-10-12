import { readFileSync } from 'node:fs';
import { describe, expect, it } from 'vitest';
import { buildCompoundFile, listCompoundFile, parseCompoundFile, parseOle2, parseDoc, parsePpt, parseXls,
	repairCompoundFile } from '../src/index.js';
import { v4Cfb } from './fixtures/v4-cfb.js';

const open = (bytes: Uint8Array) => parseOle2(bytes, { expect: 'cfb' });
const version = (bytes: Uint8Array) => new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength).getUint16(0x1a, true);
const payload = (length: number) => Uint8Array.from({ length }, (_, index) => index * 17 & 255);

function checkV4(bytes: Uint8Array): void {
	const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
	expect(version(bytes)).toBe(4); expect(view.getUint16(0x1e, true)).toBe(12);
	expect(bytes.length % 4096).toBe(0);
	expect(bytes.subarray(512, 4096).every(byte => byte === 0)).toBe(true);
	expect(parseCompoundFile(bytes, { strict: true }).warnings).toEqual([]);
}

describe('CFB v4 structural editing', () => {
	it('edits an independently generated v4 fixture without downgrading, including mini and empty entries', () => {
		const input = v4Cfb([
			{ path: ['Source', 'Nested', 'Mini'], bytes: payload(65) },
			{ path: ['Source', 'Large'], bytes: payload(8193) },
			{ path: ['Other', 'Mini'], bytes: payload(63) },
		]);
		const before = listCompoundFile(input), document = open(input);
		const source = document.storage(['Source']), mini = source.storage('Nested').stream('Mini');
		const unrelated = document.stream(['Other', 'Mini']);
		const created = document.root.createStorage('Created');
		created.createStorage('EmptyStorage'); created.createStream('EmptyStream');
		created.createStream('BoundaryMini', payload(4095)); created.createStream('BoundaryRegular', payload(4096));
		source.rename('Renamed'); source.move(['Created'], 'Moved');
		expect(mini.path).toEqual(['Created', 'Moved', 'Nested', 'Mini']);
		expect(mini.bytes).toEqual(payload(65)); expect(unrelated.bytes).toEqual(payload(63));
		mini.rename('MINI'); mini.move(['Other'], 'MovedMini');
		expect(mini.bytes).toEqual(payload(65));
		const saved = document.serialize(); checkV4(saved);
		expect(open(saved).storage(['Created', 'EmptyStorage']).children).toEqual([]);
		expect(open(saved).stream(['Created', 'EmptyStream']).bytes).toEqual(new Uint8Array());
		expect(open(saved).stream(['Created', 'BoundaryMini']).bytes).toEqual(payload(4095));
		expect(open(saved).stream(['Created', 'BoundaryRegular']).bytes).toEqual(payload(4096));
		document.move(['Created', 'Moved'], ['Source']); mini.move(['Source', 'Nested'], 'Mini');
		created.delete();
		const restored = listCompoundFile(document.serialize());
		for (const node of before.nodes) expect(restored.nodes.find(other => JSON.stringify(other.path) === JSON.stringify(node.path))).toEqual(node);
		checkV4(document.serialize());
		expect(() => created.children).toThrow(/stale-entry/);
	});

	it('grows and shrinks directory chains across the 32-entry v4 sector boundary', () => {
		const input = v4Cfb(Array.from({ length: 30 }, (_, i) => ({ path: [`S${i}`], bytes: payload(1) })));
		const document = open(input);
		expect(new DataView(input.buffer).getUint32(0x28, true)).toBe(1);
		document.root.createStorage('Added').createStream('Child', payload(65));
		expect(new DataView(document.serialize().buffer).getUint32(0x28, true)).toBe(2);
		checkV4(document.serialize());
		document.delete(['Added']);
		expect(new DataView(document.serialize().buffer).getUint32(0x28, true)).toBe(1);
		checkV4(document.serialize());
		for (let i = 0; i < 30; i++) expect(document.stream([`S${i}`]).bytes).toEqual(payload(1));
	});

	it('rebuilds multiple FAT and mini-FAT sectors and retains live stream handles', () => {
		const document = open(v4Cfb([{ path: ['Large'], bytes: payload(4 * 1024 * 1024) },
			...Array.from({ length: 20 }, (_, i) => ({ path: [`Mini${i}`], bytes: payload(4095) }))]));
		const large = document.stream(['Large']), mini = document.stream(['Mini19']);
		document.root.createStorage('Destination'); large.move(['Destination']); mini.rename('RenamedMini');
		const saved = document.serialize(), view = new DataView(saved.buffer);
		expect(view.getUint32(0x2c, true)).toBeGreaterThan(1);
		expect(view.getUint32(0x40, true)).toBeGreaterThan(1);
		checkV4(saved);
		expect(Buffer.from(large.bytes).equals(Buffer.from(payload(4 * 1024 * 1024)))).toBe(true);
		expect(mini.bytes).toEqual(payload(4095));
	});

	it('retains exact metadata and source version through list/build; default creation and repair remain v3', () => {
		const ticks = 12345678901234567n, clsid = new Uint8Array(16).fill(7);
		const input = buildCompoundFile([{ path: ['Source'], clsid, stateBits: 37, createdFileTime: ticks,
			modifiedFileTime: ticks + 1n }, { path: ['Source', 'Data'], data: payload(65) }],
			{ majorVersion: 4, rootClsid: clsid, rootStateBits: 9, rootModifiedFileTime: ticks });
		const framed = new Uint8Array(input.length + 11); framed.set(input, 7);
		const listed = listCompoundFile(framed.subarray(7, 7 + input.length));
		expect(listed.options.majorVersion).toBe(4);
		checkV4(buildCompoundFile(listed.nodes, listed.options));
		const document = open(input); document.rename(['Source'], 'Renamed');
		const after = listCompoundFile(document.serialize());
		expect(after.options).toEqual(listed.options);
		expect(after.nodes.find(node => node.path.length === 1)).toEqual({ ...listed.nodes[0], path: ['Renamed'] });
		expect(version(buildCompoundFile([]))).toBe(3);
		expect(version(repairCompoundFile(input).bytes)).toBe(3);
	});

	it('keeps malformed directory-count refusals atomic and exact no-ops clean', () => {
		const input = v4Cfb([{ path: ['Data'], bytes: payload(4096) }]);
		const document = open(input); document.rename(['Data'], 'Data');
		expect(document.serialize()).toEqual(input); expect(document.revision).toBe(0);
		new DataView(input.buffer).setUint32(0x28, 2, true);
		const malformed = open(input);
		for (const action of [() => malformed.root.createStorage('New'), () => malformed.rename(['Data'], 'Moved'),
			() => malformed.move(['Data'], ['Moved']), () => malformed.delete(['Data'])]) {
			expect(action).toThrow(/directory sector count/);
			expect(malformed.serialize()).toEqual(input); expect(malformed.revision).toBe(0); expect(malformed.dirty).toBe(false);
		}
	});

	it.each(['doc', 'ppt', 'xls'] as const)('preserves %s models and all stream bytes in a v4-rewrapped Office fixture', kind => {
		const paths = { doc: 'doc/rich-runs.doc', ppt: 'ppt/native-text.ppt', xls: 'xls/workbook-features.xls' };
		const original = new Uint8Array(readFileSync(new URL(`./fixtures/${paths[kind]}`, import.meta.url)));
		const snapshot = (bytes: Uint8Array) => {
			if (kind === 'doc') return parseDoc(bytes).paragraphs.map(p => ({ text: p.text, style: p.styleIndex,
				runs: p.runs.map(run => ({ text: run.text, bold: run.directBold, italic: run.directItalic, size: run.directFontSizePoints })) }));
			if (kind === 'ppt') return parsePpt(bytes).slides.map(slide => ({ id: slide.id, texts: slide.texts.map(text => text.text),
				shapes: slide.shapes.map(shape => ({ id: shape.id, kind: shape.kind, bounds: shape.bounds })), notes: slide.notes?.texts.map(text => text.text) }));
			return parseXls(bytes).workbook;
		};
		const baseline = snapshot(original), listed = listCompoundFile(original);
		const v4 = buildCompoundFile(listed.nodes, { ...listed.options, majorVersion: 4 });
		expect(snapshot(v4)).toEqual(baseline);
		const document = open(v4), added = document.root.createStorage('Attachments');
		added.createStream('Payload', payload(4095)); added.rename('Renamed'); added.move([], 'Moved');
		added.createStorage('Temporary').delete();
		const saved = document.serialize(); checkV4(saved);
		expect(snapshot(saved)).toEqual(baseline);
		const after = listCompoundFile(saved);
		for (const node of listed.nodes) expect(after.nodes.find(other => JSON.stringify(other.path) === JSON.stringify(node.path))).toEqual(node);
		added.delete(); checkV4(document.serialize()); expect(snapshot(document.serialize())).toEqual(baseline);
	});
});
