import { describe, expect, it } from 'vitest';
import { buildOle2 } from '../src/ole2-parser-write.js';
import { parseOle2 } from '../src/ole2-parser-read.js';
import { readCompoundFileStream, replaceCompoundFileStream } from '../src/ole2-stream-edit.js';

function nested(input: Uint8Array, streamName: string): Uint8Array {
	const out = input.slice();
	const view = new DataView(out.buffer);
	const dirSector = view.getUint32(0x30, true);
	const dirOffset = (dirSector + 1) * 512;
	const root = dirOffset;
	const stream = dirOffset + 128;
	const child = dirOffset + 256;
	const source = out.slice(stream, stream + 128);
	const putName = (at: number, value: string) => {
		for (let i = 0; i < 32; i++) view.setUint16(at + i * 2, 0, true);
		for (let i = 0; i < value.length; i++) view.setUint16(at + i * 2, value.charCodeAt(i), true);
		view.setUint16(at + 64, (value.length + 1) * 2, true);
	};
	view.setUint32(root + 76, 1, true);
	putName(stream, 'ObjectPool');
	out[stream + 66] = 1;
	view.setUint32(stream + 68, 0xffffffff, true);
	view.setUint32(stream + 72, 0xffffffff, true);
	view.setUint32(stream + 76, 2, true);
	out.set(source, child);
	putName(child, streamName);
	out[child + 66] = 2;
	view.setUint32(child + 68, 0xffffffff, true);
	view.setUint32(child + 72, 0xffffffff, true);
	view.setUint32(child + 76, 0xffffffff, true);
	return out;
}

function duplicateLeaf(input: Uint8Array): Uint8Array {
	const out = nested(input, 'Dup');
	const view = new DataView(out.buffer);
	const dirOffset = (view.getUint32(0x30, true) + 1) * 512;
	const storage = dirOffset + 640,
		leaf = dirOffset + 768;
	const setName = (at: number, name: string) => {
		for (let i = 0; i < 32; i++) view.setUint16(at + i * 2, 0, true);
		for (let i = 0; i < name.length; i++) view.setUint16(at + i * 2, name.charCodeAt(i), true);
		view.setUint16(at + 64, (name.length + 1) * 2, true);
	};
	setName(dirOffset + 128, 'Alpha');
	view.setUint32(dirOffset + 128 + 72, 5, true);
	setName(storage, 'Beta');
	out[storage + 66] = 1;
	view.setUint32(storage + 68, 0xffffffff, true);
	view.setUint32(storage + 72, 0xffffffff, true);
	view.setUint32(storage + 76, 6, true);
	setName(leaf, 'Dup');
	view.setUint32(leaf + 68, 0xffffffff, true);
	view.setUint32(leaf + 72, 0xffffffff, true);
	view.setUint32(leaf + 76, 0xffffffff, true);
	return out;
}

describe('replaceCompoundFileStream', () => {
	it('replaces a mini stream and leaves the rest of the container intact', () => {
		const original = new Uint8Array([1, 2, 3, 4, 5]);
		const file = new Uint8Array(buildOle2(new Map([['Small', original]])));
		const replacement = new Uint8Array([5, 4, 3, 2, 1]);
		const updated = replaceCompoundFileStream(file, ['Small'], replacement);
		const parsed = parseOle2(
			updated.buffer.slice(updated.byteOffset, updated.byteOffset + updated.byteLength),
		);
		expect(parsed.getStream('Small')).toEqual(replacement);
		const diffs = Array.from(updated).flatMap((byte, i) => (byte === file[i] ? [] : [i]));
		expect(diffs.every((i) => i >= 512)).toBe(true);
	});

	it('replaces a nested mini stream while preserving every other byte', () => {
		const original = new Uint8Array(5).fill(4);
		const file = nested(new Uint8Array(buildOle2(new Map([['Workbook', original]]))), 'Book');
		const replacement = new Uint8Array(5).fill(9);
		expect(readCompoundFileStream(file, ['ObjectPool', 'Book'])).toEqual(new Uint8Array(5).fill(4));
		const updated = replaceCompoundFileStream(file, ['ObjectPool', 'Book'], replacement);
		expect(updated).not.toBe(file);
		const reparsed = parseOle2(
			updated.buffer.slice(updated.byteOffset, updated.byteOffset + updated.byteLength),
		);
		expect(reparsed.getStream('Book')).toEqual(replacement);
		expect(readCompoundFileStream(updated, ['ObjectPool', 'Book'])).toEqual(replacement);
		const diffs = Array.from(updated).flatMap((byte, i) => (byte === file[i] ? [] : [i]));
		expect(diffs.length).toBeGreaterThan(0);
	});

	it('replaces a regular stream without touching directory or FAT sectors', () => {
		const original = new Uint8Array(5000).fill(0x31);
		const file = new Uint8Array(buildOle2(new Map([['Large', original]])));
		const replacement = new Uint8Array(5000).fill(0x72);
		const updated = replaceCompoundFileStream(file, ['Large'], replacement);
		const stream = parseOle2(
			updated.buffer.slice(updated.byteOffset, updated.byteOffset + updated.byteLength),
		).getStream('Large');
		expect(stream).toEqual(replacement);
		const diffs = Array.from(updated).flatMap((byte, i) => (byte === file[i] ? [] : [i]));
		const dirSector = new DataView(file.buffer).getUint32(0x30, true);
		expect(diffs.every((i) => Math.floor(i / 512) - 1 !== dirSector)).toBe(true);
	});

	it('fails closed for changed lengths and ambiguous leaf paths', () => {
		const file = new Uint8Array(buildOle2(new Map([['Small', new Uint8Array([1, 2, 3])]])));
		expect(replaceCompoundFileStream(file, ['Small'], new Uint8Array([1, 2]))).toBe(file);
		expect(replaceCompoundFileStream(file, ['missing', 'Small'], new Uint8Array([3, 2, 1]))).toBe(
			file,
		);
	});

	it('uses the full storage path to distinguish duplicate leaf names', () => {
		const file = duplicateLeaf(
			new Uint8Array(
				buildOle2(
					new Map([
						['AAAA', new Uint8Array([1, 2])],
						['ZDummy1', new Uint8Array([0])],
						['ZDummy2', new Uint8Array([0])],
						['ZDummy3', new Uint8Array([0])],
						['ZDummy4', new Uint8Array([0])],
						['ZDummy5', new Uint8Array([0, 0])],
					]),
				),
			),
		);
		expect(readCompoundFileStream(file, ['Alpha', 'Dup'])).toEqual(new Uint8Array([1, 2]));
		expect(readCompoundFileStream(file, ['Beta', 'Dup'])).toEqual(new Uint8Array([0, 0]));
		const updated = replaceCompoundFileStream(file, ['Beta', 'Dup'], new Uint8Array([8, 9]));
		expect(readCompoundFileStream(updated, ['Alpha', 'Dup'])).toEqual(new Uint8Array([1, 2]));
		expect(readCompoundFileStream(updated, ['Beta', 'Dup'])).toEqual(new Uint8Array([8, 9]));
	});

	it('rejects cyclic allocation chains and encrypted containers', () => {
		const regular = new Uint8Array(buildOle2(new Map([['Large', new Uint8Array(5000).fill(1)]])));
		const fatSector = new DataView(regular.buffer).getUint32(0x4c, true);
		const dirSector = new DataView(regular.buffer).getUint32(0x30, true);
		const streamStart = new DataView(regular.buffer, (dirSector + 1) * 512 + 128, 128).getUint32(
			116,
			true,
		);
		const cyclic = regular.slice();
		new DataView(cyclic.buffer, (fatSector + 1) * 512).setUint32(
			streamStart * 4,
			streamStart,
			true,
		);
		expect(replaceCompoundFileStream(cyclic, ['Large'], new Uint8Array(5000))).toBe(cyclic);
		expect(readCompoundFileStream(cyclic, ['Large'])).toBeUndefined();

		const encrypted = new Uint8Array(
			buildOle2(
				new Map([
					['EncryptionInfo', new Uint8Array([1])],
					['Payload', new Uint8Array([2, 3])],
				]),
			),
		);
		expect(replaceCompoundFileStream(encrypted, ['Payload'], new Uint8Array([3, 2]))).toBe(
			encrypted,
		);
	});
});
