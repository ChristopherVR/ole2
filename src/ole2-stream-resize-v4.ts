/** Preservation-safe resizing of regular streams in v4 MS-CFB containers. */
import { readCompoundFileStream, replaceCompoundFileStream } from './ole2-stream-edit.js';

const END = 0xfffffffe, FREE = 0xffffffff, FAT = 0xfffffffd;

import type { CompoundFileStreamResizeFailure, CompoundFileStreamResizeResult } from './ole2-stream-resize.js';

type Slot = { id: number; name: string; type: number; left: number; right: number; child: number; start: number; size: number; offset: number };

function follow(table: readonly number[], start: number): number[] {
	const ids: number[] = [], seen = new Set<number>();
	for (let id = start; id !== END; id = table[id]!) {
		if (id >= table.length || seen.has(id)) throw new Error('Invalid allocation chain');
		seen.add(id);
		ids.push(id);
	}
	return ids;
}

function findSlot(slots: Map<number, Slot>, path: readonly string[]): Slot {
	let parent = slots.get(0);
	if (!parent || parent.type !== 5) throw new Error('Invalid root');
	const parents = new Set<number>();
	for (let part = 0; part < path.length; part++) {
		if (parents.has(parent.id)) throw new Error('Cyclic storage hierarchy');
		parents.add(parent.id);
		const stack = [parent.child], seen = new Set<number>(), matches: Slot[] = [];
		while (stack.length) {
			const id = stack.pop()!;
			if (id === FREE) continue;
			if (seen.has(id)) throw new Error('Invalid directory tree');
			seen.add(id);
			const slot = slots.get(id);
			if (!slot) throw new Error('Missing directory slot');
			if (slot.name.toLocaleLowerCase() === path[part]!.toLocaleLowerCase()) matches.push(slot);
			stack.push(slot.left, slot.right);
		}
		if (matches.length !== 1) throw new Error('Missing or ambiguous stream');
		parent = matches[0]!;
		if (parent.type !== (part === path.length - 1 ? 2 : 1)) throw new Error('Invalid storage path');
	}
	return parent;
}

function validateHierarchy(slots: Map<number, Slot>): void {
	const owners = new Set<number>(), storages = [0];
	while (storages.length) {
		const storage = slots.get(storages.pop()!)!;
		const stack = [storage.child], names = new Set<string>();
		while (stack.length) {
			const id = stack.pop()!;
			if (id === FREE) continue;
			const slot = slots.get(id);
			if (!slot || (slot.type !== 1 && slot.type !== 2) || owners.has(id)) throw new Error('Ambiguous directory ownership');
			const name = slot.name.toLocaleLowerCase();
			if (names.has(name)) throw new Error('Duplicate sibling name');
			names.add(name);
			owners.add(id);
			stack.push(slot.left, slot.right);
			if (slot.type === 1) storages.push(id);
		}
	}
	if (owners.size !== slots.size - 1) throw new Error('Unreachable directory entries');
}

/** Replace a regular stream, allowing its payload length to change. The path
 * includes every storage name relative to the root. Existing allocation and
 * directory bytes are retained; only target start/size, affected FAT entries
 * and (if necessary) header FAT count/DIFAT slots change. New payload sectors
 * are appended; old target sectors become free without erasing their bytes.
 *
 * Scope: v4, 4096-byte sectors, header-only DIFAT (at most 109 FAT
 * sectors), cutoff 4096 or the existing cutoff-0 compatibility layout.
 * Resizing mini streams and transitions between mini/regular allocations are
 * explicitly unsupported. Same-length edits still use the existing validator.
 * Encrypted or overlapping allocations fail safely. Failure returns the exact
 * original input reference and a reason. No input buffer is mutated.
 *
 * This edits the CFB container only. A format-specific caller must also repair
 * offsets and lengths inside its DOC/XLS/PPT stream payload.
 * Reference: [MS-CFB] sections 2.2, 2.3 and 2.6.
 */
export function resizeCompoundFileStreamV4(
	input: Uint8Array,
	path: readonly string[],
	replacement: Uint8Array,
): CompoundFileStreamResizeResult {
	const fail = (reason: CompoundFileStreamResizeFailure): CompoundFileStreamResizeResult => ({ ok: false, bytes: input, reason });
	try {
		if (input.length < 512 || !path.length || path.some((part) => !part)) return fail('invalid-or-missing-stream');
		const view = new DataView(input.buffer, input.byteOffset, input.byteLength);
		const read = (offset: number) => view.getUint32(offset, true);
		const version = view.getUint16(0x1a, true);
		if (version !== 4) return fail('unsupported-version');
		const SECTOR = 4096, FAT_ENTRIES = SECTOR / 4, DIRECTORY_ENTRIES = SECTOR / 128;
		if (read(0x48) !== 0 || read(0x2c) > 109) return fail('unsupported-difat');
		const cutoff = read(0x38);
		if (cutoff !== 0 && cutoff !== 4096) return fail('unsupported-layout');
		const original = readCompoundFileStream(input, path);
		if (!original) return fail('invalid-or-missing-stream');
		const name = path[path.length - 1]!.toLocaleLowerCase();
		if (name === 'worddocument' && original.length >= 12 && ((original[10]! | (original[11]! << 8)) & 0x8100)) return fail('unsafe-edit');
		if (name === 'book' || name === 'workbook') {
			for (let offset = 0; offset + 4 <= original.length;) {
				const record = original[offset]! | (original[offset + 1]! << 8);
				const length = original[offset + 2]! | (original[offset + 3]! << 8);
				if (record === 0x002f) return fail('unsafe-edit');
				if (record === 0x000a || offset + 4 + length > original.length) break;
				offset += 4 + length;
			}
		}
		// Reuse the existing global allocation-alias and encryption safety gate.
		const validated = replaceCompoundFileStream(input, path, original);
		if (validated === input) return fail('unsafe-edit');
		const sameLength = replacement.length === original.length;
		if (!sameLength && original.length < cutoff) return fail('unsupported-mini-stream');
		if (!sameLength && replacement.length < cutoff) return fail('unsupported-mini-transition');

		const sectorCount = input.length / SECTOR - 1, fatCount = read(0x2c);
		if (sectorCount > fatCount * FAT_ENTRIES) return fail('unsafe-edit');
		const fatIds = Array.from({ length: fatCount }, (_, i) => read(0x4c + 4 * i));
		const fat: number[] = [];
		for (const id of fatIds) {
			for (let i = 0; i < FAT_ENTRIES && fat.length < sectorCount; i++) fat.push(read((id + 1) * SECTOR + i * 4));
		}
		if (fatIds.some((id) => fat[id] !== FAT)) return fail('unsafe-edit');
		const slots = new Map<number, Slot>();
		const directoryChain = follow(fat, read(0x30));
		if (read(0x28) !== directoryChain.length) return fail('unsafe-edit');
		for (const [block, sector] of directoryChain.entries()) {
			for (let entry = 0; entry < DIRECTORY_ENTRIES; entry++) {
				const offset = (sector + 1) * SECTOR + entry * 128, type = view.getUint8(offset + 66);
				if (!type) continue;
				const nameLength = view.getUint16(offset + 64, true);
				if (![1, 2, 5].includes(type) || nameLength < 2 || nameLength > 64 || nameLength % 2 || view.getUint16(offset + nameLength - 2, true)) return fail('unsafe-edit');
				let name = '';
				for (let i = 0; i < nameLength - 2; i += 2) name += String.fromCharCode(view.getUint16(offset + i, true));
				// Existing safe stream reader rejects 64-bit lengths. Keep that
				// refusal explicit here rather than truncating v4 directory sizes.
				if (read(offset + 124) !== 0) return fail('unsafe-edit');
				const id = block * DIRECTORY_ENTRIES + entry;
				slots.set(id, { id, name, type, left: read(offset + 68), right: read(offset + 72), child: read(offset + 76), start: read(offset + 116), size: read(offset + 120), offset });
			}
		}
		if ([...slots.values()].filter((slot) => slot.type === 5).length !== 1) return fail('unsafe-edit');
		const target = findSlot(slots, path);
		validateHierarchy(slots);
		if (sameLength) {
			const bytes = replaceCompoundFileStream(input, path, replacement);
			return bytes === input ? fail('unsafe-edit') : { ok: true, bytes };
		}
		const oldChain = follow(fat, target.start);
		if (target.size !== original.length || oldChain.length !== Math.ceil(target.size / SECTOR)) return fail('unsafe-edit');
		const payloadSectors = Math.ceil(replacement.length / SECTOR);
		let newFatCount = fatCount;
		while (sectorCount + payloadSectors + newFatCount - fatCount > newFatCount * FAT_ENTRIES) newFatCount++;
		if (newFatCount > 109) return fail('allocation-limit');
		const totalSectors = sectorCount + payloadSectors + newFatCount - fatCount;
		const output = new Uint8Array((totalSectors + 1) * SECTOR);
		output.set(input);
		output.set(replacement, (sectorCount + 1) * SECTOR);
		const outView = new DataView(output.buffer);
		for (let i = fatCount; i < newFatCount; i++) {
			const id = sectorCount + payloadSectors + i - fatCount;
			fatIds.push(id);
			output.fill(0xff, (id + 1) * SECTOR, (id + 2) * SECTOR);
			outView.setUint32(0x4c + i * 4, id, true);
		}
		outView.setUint32(0x2c, newFatCount, true);
		const writeFat = (id: number, value: number) => outView.setUint32((fatIds[Math.floor(id / FAT_ENTRIES)]! + 1) * SECTOR + (id % FAT_ENTRIES) * 4, value, true);
		for (const id of oldChain) writeFat(id, FREE);
		for (let i = 0; i < payloadSectors; i++) writeFat(sectorCount + i, i + 1 === payloadSectors ? END : sectorCount + i + 1);
		for (const id of fatIds) writeFat(id, FAT);
		outView.setUint32(target.offset + 116, payloadSectors ? sectorCount : END, true);
		outView.setUint32(target.offset + 120, replacement.length, true);
		return { ok: true, bytes: output };
	} catch {
		return fail('unsafe-edit');
	}
}
