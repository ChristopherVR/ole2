import { readFileSync } from 'node:fs';
import { describe, expect, it } from 'vitest';
import { tryWriteOleDocParagraphEdit, readOleDocParagraphs } from '../src/ole-document-doc-editor.js';
import { DocCfbRewriteError, unwrapDocBytes } from '../src/ole-document-doc-cfb.js';
import { parseOle2 } from '../src/ole2-parser-read.js';
import { readCompoundFileStream } from '../src/ole2-stream-edit.js';
import { buildOle2 } from '../src/ole2-parser-write.js';
import { nestedCfb } from './fixtures/nested-cfb.js';

const load = () => new Uint8Array(readFileSync(new URL('./fixtures/ole-word-97.doc', import.meta.url)));

function directory(bytes: Uint8Array): Uint8Array {
	const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
	const sector = view.getUint32(0x30, true);
	return bytes.slice((sector + 1) * 512, (sector + 1) * 512 + 2048);
}

function nestedDoc(): Uint8Array {
	const original = load();
	const ole = parseOle2(original.buffer as ArrayBuffer);
	const streams = ole.entries.filter((entry) => entry.type === 2).map((entry) => ({ path: [entry.name], bytes: ole.getStream(entry.name)! }));
	streams.push(
		{ path: ['ObjectPool', 'Fixture', 'Private'], bytes: new Uint8Array([7, 8, 9]) },
		{ path: ['Custom', 'Private'], bytes: new Uint8Array([42, 43, 44]) },
		{ path: ['Custom', 'WordDocument'], bytes: new Uint8Array([1, 2, 3, 4]) },
	);
	const bytes = nestedCfb(streams);
	const view = new DataView(bytes.buffer);
	const start = (view.getUint32(0x30, true) + 1) * 512;
	const entries = parseOle2(bytes.buffer as ArrayBuffer).entries;
	for (let id = 0; id < entries.length; id++) {
		const entry = entries[id]!;
		if (entry.type !== 1) continue;
		const off = start + id * 128;
		bytes.fill(id + 1, off + 80, off + 96); // opaque storage CLSID
		view.setUint32(off + 96, 0x70000000 + id, true); // state bits
		view.setBigUint64(off + 100, 133000000000000000n + BigInt(id), true);
		view.setBigUint64(off + 108, 134000000000000000n + BigInt(id), true);
	}
	return bytes;
}

describe('preserving DOC compound stream growth', () => {
	it('grows WordDocument/table without flattening nested streams or rewriting metadata', () => {
		const input = nestedDoc();
		const beforeText = readOleDocParagraphs(input)!;
		const before = directory(input);
		const replacement = 'Nested DOC preservation: ' + 'larger paragraph '.repeat(500);
		const result = tryWriteOleDocParagraphEdit(input, 2, replacement);
		expect(result.status).toBe('edited');
		if (result.status !== 'edited') throw new Error(result.reason);
		expect(result.strategy).toBe('piece-append');
		expect(readOleDocParagraphs(result.bytes)).toStrictEqual([beforeText[0], beforeText[1], replacement, beforeText[3]]);
		for (const [path, expected] of [
			[['ObjectPool', 'Fixture', 'Private'], [7, 8, 9]],
			[['Custom', 'Private'], [42, 43, 44]],
			[['Custom', 'WordDocument'], [1, 2, 3, 4]],
		] as const) expect(readCompoundFileStream(result.bytes, path)).toStrictEqual(new Uint8Array(expected));
		const after = directory(result.bytes);
		const entries = parseOle2(input.buffer as ArrayBuffer).entries;
		const table = unwrapDocBytes(input)!.tableStreamName;
		for (let id = 0; id < entries.length; id++) {
			const entry = entries[id]!;
			// Only the root-level DOC streams' allocation start/size fields change.
			const target = entry.type === 2 && (entry.name === table || entry.name === 'WordDocument') && entry.size >= 4096;
			const end = id * 128 + (target ? 116 : 128);
			expect(after.slice(id * 128, end)).toStrictEqual(before.slice(id * 128, end));
		}
		const originalOle = parseOle2(input.buffer as ArrayBuffer);
		for (const entry of originalOle.entries) {
			if (entry.type === 2 && !['WordDocument', table].includes(entry.name))
				expect(readCompoundFileStream(result.bytes, [entry.name])).toStrictEqual(readCompoundFileStream(input, [entry.name]));
		}
	});

	it('refuses unsupported mini-stream growth without flattening the original container', () => {
		const doc = unwrapDocBytes(load())!;
		const miniTable = doc.tableBytes.slice(0, 128);
		const input = new Uint8Array(buildOle2(new Map([
			['WordDocument', doc.wordDocBytes], [doc.tableStreamName, miniTable],
		])));
		const parsed = unwrapDocBytes(input)!;
		expect(parsed.canRewrite).toBe(false);
		expect(() => parsed.rewrap(doc.wordDocBytes, new Uint8Array(5000))).toThrow(DocCfbRewriteError);
		expect(readCompoundFileStream(input, [doc.tableStreamName])).toStrictEqual(miniTable);
	});
});
