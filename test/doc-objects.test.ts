import { readFileSync } from 'node:fs';
import { deflateSync, deflateRawSync } from 'node:zlib';
import { describe, expect, it } from 'vitest';
import { DocInflateError, DocResourceLimitError, inflateRaw, inflateZlib, parseFieldInstruction, readCompObjProgId, readDocObjects } from '../src/index.js';

const fixture = (path: string) => new Uint8Array(readFileSync(new URL(`./fixtures/${path}`, import.meta.url)));
const png = fixture('doc/objects-input-a.png');
const jpg = fixture('doc/objects-input-b.jpg');
const doc = fixture('doc/objects.doc');
const objects = readDocObjects(doc)!;
const same = (a: Uint8Array, b: Uint8Array) => Buffer.compare(a, b) === 0;

describe('readDocObjects', () => {
	it('returns undefined for non-DOC input', () => {
		expect(readDocObjects(new Uint8Array(64))).toBeUndefined();
		expect(readDocObjects(fixture('xls/workbook-mulrk.xls'))).toBeUndefined();
	});

	it('reads story ranges and reports no warnings for a clean file', () => {
		expect(objects.stories.map((s) => s.story)).toEqual(['main', 'textboxes']);
		expect(objects.warnings).toEqual([]);
	});

	it('decodes inline pictures byte-for-byte with their display size', () => {
		const [first, second] = objects.inlinePictures;
		expect(first!.blip!.format).toBe('png');
		expect(same(first!.blip!.data, png)).toBe(true);
		expect([first!.blip!.widthPx, first!.blip!.heightPx]).toEqual([64, 48]);
		expect([first!.dxaGoal, first!.dyaGoal, first!.mx, first!.my]).toEqual([960, 720, 2000, 2000]);
		expect([first!.widthTwips, first!.heightTwips]).toEqual([1920, 1440]);
		expect(first!.alt).toBe('inline png');
		expect(second!.blip!.format).toBe('jpeg');
		expect(same(second!.blip!.data, jpg)).toBe(true);
		expect([second!.widthTwips, second!.heightTwips]).toEqual([1200, 600]);
		expect(objects.stories[0]!.cpEnd).toBeGreaterThan(second!.cp);
	});

	it('does not count an embedded object preview as an inline picture', () => {
		const preview = objects.oleObjects[0]!.preview!;
		expect(objects.inlinePictures.some((p) => p.cp === preview.cp)).toBe(false);
	});

	it('reads the document BLIP store, including out-of-line BLIPs', () => {
		expect(objects.blips.map((b) => b.format)).toEqual(['png', 'jpeg']);
		expect(same(objects.blips[0]!.data, png)).toBe(true);
		expect(objects.blips[0]!.delayed).toBe(true);
	});

	it('joins FSPA anchors with floating pictures and shapes', () => {
		const byName = new Map(objects.floatingShapes.map((s) => [s.name, s]));
		const pic = byName.get('FloatPng')!;
		expect(pic.isPicture).toBe(true);
		expect(pic.wrap).toBe('square');
		expect(pic.relativeX).toBe('page');
		expect(pic.relativeY).toBe('paragraph');
		expect([pic.leftTwips, pic.topTwips, pic.widthTwips, pic.heightTwips]).toEqual([6000, 2400, 2000, 1500]);
		expect(same(pic.blip!.data, png)).toBe(true);
		const front = byName.get('FloatJpg')!;
		expect(front.wrap).toBe('none');
		expect(front.widthTwips).toBe(1600);
		expect(same(front.blip!.data, jpg)).toBe(true);
		const rect = byName.get('Rect')!;
		expect(rect.shapeType).toBe(1);
		expect(rect.isPicture).toBe(false);
		expect(rect.fill).toMatchObject({ enabled: true, colorRgb: 0xff0000 });
		expect(rect.line).toMatchObject({ enabled: true, colorRgb: 0x00ff00, widthEmu: 38100 });
		expect([rect.widthTwips, rect.heightTwips]).toEqual([2400, 1200]);
		const box = byName.get('Tbox')!;
		expect(box.shapeType).toBe(202);
		expect(box.textboxId).toBe(0x10000);
	});

	it('decodes hyperlink fields', () => {
		const links = objects.fields.filter((f) => f.kind === 'hyperlink');
		expect(links).toHaveLength(2);
		expect(links[0]).toMatchObject({ url: 'https://example.com/path?x=1', tooltip: 'Example tip', result: 'ext link', flt: 88 });
		expect(links[1]).toMatchObject({ anchor: 'BM_Target', tooltip: 'Local tip', result: 'to bookmark' });
		expect(links[0]!.cpSeparator).toBeGreaterThan(links[0]!.cpBegin);
		expect(links[0]!.resultRange).toEqual({ cpStart: links[0]!.cpSeparator! + 1, cpEnd: links[0]!.cpEnd });
	});

	it('decodes REF, PAGEREF, PAGE, NUMPAGES, DATE, SEQ, MERGEFIELD and INCLUDEPICTURE', () => {
		const of = (kind: string) => objects.fields.filter((f) => f.kind === kind);
		expect(of('ref')[0]).toMatchObject({ bookmark: 'BM_Target', switches: { h: true } });
		expect(of('pageref').map((f) => (f as { bookmark: string }).bookmark)).toEqual(['BM_Target', '_Toc244260242']);
		expect(of('page')[0]).toMatchObject({ instruction: 'PAGE', result: '1', flt: 33 });
		expect(of('numpages')[0]).toMatchObject({ instruction: 'NUMPAGES', flt: 26 });
		expect(of('date')[0]).toMatchObject({ format: 'yyyy-MM-dd', result: '2026-10-30' });
		expect(of('seq')[0]).toMatchObject({ identifier: 'Figure', result: '1' });
		expect(of('mergefield')[0]).toMatchObject({ name: 'CustomerName', result: '\u00abCustomerName\u00bb' });
		expect(of('includepicture')[0]).toMatchObject({ path: 'C:\\nonexistent\\pic.png', linked: false });
	});

	it('nests fields and decodes TOC', () => {
		const toc = objects.fields.findIndex((f) => f.kind === 'toc');
		expect(objects.fields[toc]).toMatchObject({ outlineLevels: '1-1', hyperlinks: false, depth: 0, parent: -1 });
		const nested = objects.fields.filter((f) => f.parent === toc);
		expect(nested).toHaveLength(1);
		expect(nested[0]).toMatchObject({ kind: 'pageref', depth: 1 });
		expect(nested[0]!.cpBegin).toBeGreaterThan(objects.fields[toc]!.cpSeparator!);
		expect(nested[0]!.cpEnd).toBeLessThan(objects.fields[toc]!.cpEnd);
	});

	it('reads bookmark names and ranges', () => {
		const named = new Map(objects.bookmarks.map((b) => [b.name, b]));
		expect(named.get('BM_Target')).toMatchObject({ cpStart: 180, cpEnd: 292, hidden: false, story: 'main' });
		expect(named.get('BM_Point')).toMatchObject({ cpStart: 116, cpEnd: 159 });
		expect(named.get('_Toc244260242')!.hidden).toBe(true);
	});

	it('reads the embedded object storage, ProgID, class id and EMF preview', () => {
		expect(objects.oleObjects).toHaveLength(1);
		const ole = objects.oleObjects[0]!;
		expect(ole).toMatchObject({ storageName: '_1854872982', progId: 'Excel.Sheet.8', clsid: '00020820-0000-0000-c000-000000000046', linked: false });
		expect(ole.streams).toContain('Workbook');
		const blip = ole.preview!.blip!;
		expect(blip.format).toBe('emf');
		// EMF header: record type 1, then the " EMF" signature at offset 40.
		expect(Array.from(blip.data.subarray(0, 4))).toEqual([1, 0, 0, 0]);
		expect(String.fromCharCode(...blip.data.subarray(40, 44))).toBe(' EMF');
		expect(blip.metafile!.sizeEmu[0]).toBeGreaterThan(0);
	});

	it('enforces processing limits', () => {
		expect(() => readDocObjects(doc, { maxRecords: 1 })).toThrow(DocResourceLimitError);
		expect(() => readDocObjects(doc, { maxCharacters: 10 })).toThrow(DocResourceLimitError);
		const small = readDocObjects(doc, { maxBlipBytes: 100 })!;
		expect(small.blips).toEqual([]);
		expect(small.inlinePictures[0]!.blip).toBeUndefined();
		expect(() => readDocObjects(doc, { maxRecords: -1 })).toThrow();
	});

	it('survives hostile byte corruption without unexpected errors', () => {
		let seed = 12345;
		const rnd = () => (seed = (seed * 1103515245 + 12345) >>> 0) / 4294967296;
		for (let round = 0; round < 60; round++) {
			const copy = doc.slice();
			for (let i = 0; i < 40; i++) copy[Math.floor(rnd() * copy.length)] = Math.floor(rnd() * 256);
			try { readDocObjects(copy, { maxBlipBytes: 1 << 20, maxTotalBlipBytes: 1 << 22 }); }
			catch (error) { expect(error).toBeInstanceOf(DocResourceLimitError); }
		}
	});
});

describe('parseFieldInstruction', () => {
	it('splits name, arguments and switches with quoting', () => {
		expect(parseFieldInstruction('HYPERLINK "http://a/b" \\o "Tip text" \\l "Here" \\n')).toEqual({
			name: 'HYPERLINK', args: ['http://a/b'], switches: { o: 'Tip text', l: 'Here', n: true },
		});
		expect(parseFieldInstruction(' SEQ Fig \\* ARABIC \\* MERGEFORMAT \\r 5').switches).toEqual({ '*': 'ARABIC,MERGEFORMAT', r: '5' });
		expect(parseFieldInstruction('INCLUDEPICTURE "C:\\\\x\\\\y.png" \\d').args).toEqual(['C:\\x\\y.png']);
		expect(parseFieldInstruction('')).toEqual({ name: '', args: [], switches: {} });
	});
});

describe('inflate', () => {
	const sample = new Uint8Array(Array.from({ length: 5000 }, (_, i) => (i * 7 + (i >> 5)) & 0xff));
	it('inflates zlib and raw DEFLATE streams', () => {
		expect(same(inflateZlib(deflateSync(sample), 10_000), sample)).toBe(true);
		expect(same(inflateRaw(deflateRawSync(sample, { level: 0 }), 0, 10_000), sample)).toBe(true);
		expect(same(inflateRaw(deflateRawSync(sample, { strategy: 4 }), 0, 10_000), sample)).toBe(true);
	});
	it('refuses oversize, truncated and corrupt streams', () => {
		expect(() => inflateZlib(deflateSync(new Uint8Array(100_000)), 1000)).toThrow(DocInflateError);
		expect(() => inflateZlib(deflateSync(sample).subarray(0, 40), 10_000)).toThrow(DocInflateError);
		expect(() => inflateZlib(new Uint8Array([1, 2, 3]), 10)).toThrow(DocInflateError);
	});
});

describe('readCompObjProgId', () => {
	it('reads the ProgID after the user type and clipboard format', () => {
		const str = (s: string) => {
			const b = new Uint8Array(4 + s.length + 1);
			new DataView(b.buffer).setUint32(0, s.length + 1, true);
			for (let i = 0; i < s.length; i++) b[4 + i] = s.charCodeAt(i);
			return b;
		};
		const parts = [new Uint8Array(28), str('Type'), new Uint8Array([0xff, 0xff, 0xff, 0xff, 5, 0, 0, 0]), str('Prog.Id.1')];
		const all = new Uint8Array(parts.reduce((n, p) => n + p.length, 0));
		let at = 0;
		for (const p of parts) { all.set(p, at); at += p.length; }
		expect(readCompObjProgId(all)).toBe('Prog.Id.1');
		expect(readCompObjProgId(all.subarray(0, 40))).toBeUndefined();
	});
});
