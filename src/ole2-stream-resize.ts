/** Preservation-safe resizing of regular streams in v3 MS-CFB containers. */
import { readCompoundFileStream, replaceCompoundFileStream } from './ole2-stream-edit.js';

const END = 0xfffffffe, FREE = 0xffffffff, FAT = 0xfffffffd;
const SECTOR = 512, FAT_ENTRIES = SECTOR / 4;

export type CompoundFileStreamResizeFailure =
	| 'invalid-or-missing-stream'
	| 'unsafe-edit'
	| 'unsupported-version'
	| 'unsupported-difat'
	| 'unsupported-layout'
	| 'unsupported-mini-stream'
	| 'unsupported-mini-transition'
	| 'allocation-limit';

export type CompoundFileStreamResizeResult =
	| { ok: true; bytes: Uint8Array }
	| { ok: false; bytes: Uint8Array; reason: CompoundFileStreamResizeFailure };

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

/** Replace a regular stream, allowing its payload length to change. The path
 * includes every storage name relative to the root. Existing allocation and
 * directory bytes are retained; only target start/size, affected FAT entries
 * and (if necessary) header FAT count/DIFAT slots change. New payload sectors
 * are appended; old target sectors become free without erasing their bytes.
 *
 * Initial scope: v3, 512-byte sectors, header-only DIFAT (at most 109 FAT
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
export function resizeCompoundFileStream(
	input: Uint8Array,
	path: readonly string[],
	replacement: Uint8Array,
): CompoundFileStreamResizeResult {
	const fail = (reason: CompoundFileStreamResizeFailure): CompoundFileStreamResizeResult => ({ ok: false, bytes: input, reason });
	try {
		if (input.length < SECTOR || !path.length || path.some((part) => !part)) return fail('invalid-or-missing-stream');
		const view = new DataView(input.buffer, input.byteOffset, input.byteLength);
		const read = (offset: number) => view.getUint32(offset, true);
		if (view.getUint16(0x1a, true) !== 3) return fail('unsupported-version');
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
		if (replacement.length === original.length) {
			const bytes = replaceCompoundFileStream(input, path, replacement);
			return bytes === input ? fail('unsafe-edit') : { ok: true, bytes };
		}
		if (original.length < cutoff) return fail('unsupported-mini-stream');
		if (replacement.length < cutoff) return fail('unsupported-mini-transition');

		const sectorCount = input.length / SECTOR - 1, fatCount = read(0x2c);
		if (sectorCount > fatCount * FAT_ENTRIES) return fail('unsafe-edit');
		const fatIds = Array.from({ length: fatCount }, (_, i) => read(0x4c + 4 * i));
		const fat: number[] = [];
		for (const id of fatIds) {
			for (let i = 0; i < FAT_ENTRIES && fat.length < sectorCount; i++) fat.push(read((id + 1) * SECTOR + i * 4));
		}
		if (fatIds.some((id) => fat[id] !== FAT)) return fail('unsafe-edit');
		const slots = new Map<number, Slot>();
		for (const [block, sector] of follow(fat, read(0x30)).entries()) {
			for (let entry = 0; entry < 4; entry++) {
				const offset = (sector + 1) * SECTOR + entry * 128, type = view.getUint8(offset + 66);
				if (!type) continue;
				const nameLength = view.getUint16(offset + 64, true);
				if (![1, 2, 5].includes(type) || nameLength < 2 || nameLength > 64 || nameLength % 2 || view.getUint16(offset + nameLength - 2, true)) return fail('unsafe-edit');
				let name = '';
				for (let i = 0; i < nameLength - 2; i += 2) name += String.fromCharCode(view.getUint16(offset + i, true));
				const id = block * 4 + entry;
				slots.set(id, { id, name, type, left: read(offset + 68), right: read(offset + 72), child: read(offset + 76), start: read(offset + 116), size: read(offset + 120), offset });
			}
		}
		if ([...slots.values()].filter((slot) => slot.type === 5).length !== 1) return fail('unsafe-edit');
		const target = findSlot(slots, path), oldChain = follow(fat, target.start);
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
