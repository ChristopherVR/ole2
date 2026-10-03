import { describe, expect, it } from 'vitest';
import { buildPersistDirectory, parsePersistDirectoryAtom, parseUserEditAtom, readPersistRecord } from '../src/ppt/persist-directory.js';
import { PptParseError } from '../src/legacy-ppt-record-stream.js';
import { buildPptFile } from '../src/legacy-ppt-writer.js';
import { parseOle2 } from '../src/ole2-parser-read.js';
import { buildPersistDirectoryAtom, buildUserEditAtom } from '../src/ppt/writer/persist-writer.js';

function viewOf(bytes: Uint8Array): DataView {
	return new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
}

// Two valid incremental saves: the newer Document replaces id 1 and adds id 2.
function history(): DataView {
	const bytes = new Uint8Array(136);
	const view = viewOf(bytes);
	const object = (offset: number) => { view.setUint16(offset, 0xf, true); view.setUint16(offset + 2, 1000, true); };
	object(0);
	bytes.set(buildPersistDirectoryAtom([[1, 0]]), 8);
	bytes.set(buildUserEditAtom({ offsetPersistDirectory: 8, docPersistIdRef: 1, maxPersistWritten: 1, lastSlideId: 256 }), 24);
	object(60);
	object(68);
	bytes.set(buildPersistDirectoryAtom([[1, 60], [2, 68]]), 76);
	bytes.set(buildUserEditAtom({ offsetPersistDirectory: 76, docPersistIdRef: 1, maxPersistWritten: 2, lastSlideId: 256 }), 96);
	view.setUint32(96 + 8 + 8, 24, true);
	return view;
}

describe('bounded PowerPoint persist directory parsing', () => {
	it('uses newest entries across incremental saves', () => {
		const view = history();
		const chain = buildPersistDirectory(view, 96);
		expect([...chain.directory]).toEqual([[1, 60], [2, 68]]);
		expect(readPersistRecord(view, chain.directory, 1)?.headerOffset).toBe(60);
		expect(readPersistRecord(view, chain.directory, 3)).toBeUndefined();
	});

	it('accepts the shared writer output (structural check, not native fidelity)', async () => {
		const bytes = await buildPptFile({ widthEmu: 9144000, heightEmu: 5143500, slides: [{ shapes: [] }], pictures: [] });
		const file = parseOle2(bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength) as ArrayBuffer);
		const user = viewOf(file.getStream('Current User')!);
		const doc = viewOf(file.getStream('PowerPoint Document')!);
		const chain = buildPersistDirectory(doc, user.getUint32(16, true));
		expect(readPersistRecord(doc, chain.directory, 1)?.recType).toBe(1000);
	});

	it('rejects truncated atom payloads with PptParseError', () => {
		const view = history();
		expect(() => parseUserEditAtom(new DataView(view.buffer, 0, 40), 24)).toThrow(PptParseError);
		expect(() => parsePersistDirectoryAtom(new DataView(view.buffer, 0, 20), 8)).toThrow(PptParseError);
	});

	it('rejects incomplete runs, trailing bytes, zero counts, duplicate ids and overflow', () => {
		for (const mutate of [
			(v: DataView) => v.setUint32(12, 7, true), // trailing partial bytes
			(v: DataView) => v.setUint32(16, (2 << 20) | 1, true), // only one offset
			(v: DataView) => v.setUint32(16, 1, true), // zero count
			(v: DataView) => v.setUint32(16, (1 << 20) | 0xfffff, true), // invalid first id
			(v: DataView) => v.setUint32(16, (3 << 20) | 0xffffe, true),
		]) {
			const view = history(); mutate(view);
			expect(() => parsePersistDirectoryAtom(view, 8)).toThrow(PptParseError);
		}
		const bytes = new Uint8Array(32);
		bytes.set(buildPersistDirectoryAtom([[1, 0], [1, 0]]), 8);
		expect(() => parsePersistDirectoryAtom(viewOf(bytes), 8)).toThrow(PptParseError);
	});

	it('rejects invalid save ordering, cycles and object offsets', () => {
		for (const mutate of [
			(v: DataView) => v.setUint32(112, 96, true), // self-reference
			(v: DataView) => v.setUint32(112, 100, true), // forward-reference
			(v: DataView) => v.setUint32(116, 96, true), // directory at current edit
			(v: DataView) => v.setUint32(116, 8, true), // directory predates previous edit
			(v: DataView) => v.setUint32(88, 0, true), // live object predates previous edit
			(v: DataView) => v.setUint32(88, 76, true), // object inside directory
		]) {
			const view = history(); mutate(view);
			expect(() => buildPersistDirectory(view, 96)).toThrow(PptParseError);
		}
	});

	it('enforces configurable edit and aggregate entry budgets', () => {
		expect(() => buildPersistDirectory(history(), 96, { maxEdits: 1 })).toThrow(PptParseError);
		expect(() => buildPersistDirectory(history(), 96, { maxEntries: 2 })).toThrow(PptParseError);
		expect(() => buildPersistDirectory(history(), 96, { maxEdits: NaN })).toThrow(PptParseError);
	});

	it('bounds resolved record payloads and rejects invalid header/version fields', () => {
		const view = history();
		view.setUint32(64, 0xffffffff, true);
		expect(() => readPersistRecord(view, new Map([[1, 60]]), 1)).toThrow(PptParseError);
		for (const index of [24, 24 + 8 + 7]) {
			const bad = history(); bad.setUint8(index, 1);
			expect(() => parseUserEditAtom(bad, 24)).toThrow(PptParseError);
		}
	});
});
