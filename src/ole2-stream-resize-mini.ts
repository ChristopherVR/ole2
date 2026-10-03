/** Mini-stream resizing preserves existing mini IDs, root bytes and hierarchy. */
import { appendCompoundAllocations, followCompoundAllocation } from './ole2-stream-resize-allocation.js';

const END = 0xfffffffe, FREE = 0xffffffff, SECTOR = 512, MINI = 64;
type Allocation = { start: number; size: number; offset: number };

export function resizeMiniCompoundStream(
	input: Uint8Array,
	fat: readonly number[],
	fatIds: readonly number[],
	root: Allocation,
	target: Allocation,
	replacement: Uint8Array,
	cutoff: number,
): Uint8Array | undefined {
	const inputView = new DataView(input.buffer, input.byteOffset, input.byteLength);
	const miniFatCount = inputView.getUint32(0x40, true);
	const miniFatSectors = miniFatCount ? followCompoundAllocation(fat, inputView.getUint32(0x3c, true)) : [];
	if (miniFatSectors.length !== miniFatCount) throw new Error('Invalid mini FAT count');
	const rootSectors = root.size ? followCompoundAllocation(fat, root.start) : [];
	if (rootSectors.length !== Math.ceil(root.size / SECTOR)) throw new Error('Invalid root mini stream');
	const miniCapacity = Math.floor(root.size / MINI), miniFat: number[] = [];
	for (const id of miniFatSectors) {
		for (let slot = 0; slot < SECTOR / 4 && miniFat.length < miniCapacity; slot++) miniFat.push(inputView.getUint32((id + 1) * SECTOR + slot * 4, true));
	}
	const wasMini = target.size < cutoff;
	const oldChain = followCompoundAllocation(wasMini ? miniFat : fat, target.start);
	if (oldChain.length !== Math.ceil(target.size / (wasMini ? MINI : SECTOR))) throw new Error('Invalid stream allocation');
	if (!replacement.length) {
		// Deletion needs only metadata updates, so it also works when the file
		// has no remaining header-FAT capacity for appended root allocations.
		const bytes = input.slice(), view = new DataView(bytes.buffer);
		for (const id of oldChain) {
			const table = wasMini ? miniFatSectors : fatIds;
			view.setUint32((table[Math.floor(id / 128)]! + 1) * SECTOR + (id % 128) * 4, FREE, true);
		}
		view.setUint32(target.offset + 116, END, true);
		view.setUint32(target.offset + 120, 0, true);
		return bytes;
	}
	if (replacement.length >= cutoff) {
		// Mini -> regular: existing root bytes and root allocation stay intact.
		const output = appendCompoundAllocations(input, fatIds, [replacement]);
		if (!output) return undefined;
		for (const id of oldChain) output.view.setUint32((miniFatSectors[Math.floor(id / 128)]! + 1) * SECTOR + (id % 128) * 4, FREE, true);
		output.view.setUint32(target.offset + 116, output.chains[0]![0] ?? END, true);
		output.view.setUint32(target.offset + 120, replacement.length, true);
		return output.bytes;
	}

	// Append mini IDs after the full original root capacity; do not reuse free
	// root bytes, which may contain opaque application data worth retaining.
	const firstMini = Math.ceil(root.size / MINI), count = Math.ceil(replacement.length / MINI);
	const newRootSize = count ? (firstMini + count) * MINI : root.size;
	const newRoot = new Uint8Array(newRootSize);
	rootSectors.forEach((id, i) => newRoot.set(input.subarray((id + 1) * SECTOR, (id + 1) * SECTOR + Math.min(SECTOR, root.size - i * SECTOR)), i * SECTOR));
	if (count) newRoot.set(replacement, firstMini * MINI);
	const newMiniFatCount = Math.max(miniFatCount, Math.ceil((firstMini + count) / 128));
	const newMiniFat = new Uint8Array(newMiniFatCount * SECTOR).fill(0xff);
	miniFatSectors.forEach((id, i) => newMiniFat.set(input.subarray((id + 1) * SECTOR, (id + 2) * SECTOR), i * SECTOR));
	const miniView = new DataView(newMiniFat.buffer);
	if (wasMini) for (const id of oldChain) miniView.setUint32(id * 4, FREE, true);
	for (let i = 0; i < count; i++) miniView.setUint32((firstMini + i) * 4, i + 1 === count ? END : firstMini + i + 1, true);
	const output = appendCompoundAllocations(input, fatIds, [newRoot, newMiniFat]);
	if (!output) return undefined;
	// Only root, MiniFAT and a regular target being migrated are released.
	for (const id of [...rootSectors, ...miniFatSectors, ...(wasMini ? [] : oldChain)]) output.writeFat(id, FREE);
	output.view.setUint32(root.offset + 116, output.chains[0]![0] ?? END, true);
	output.view.setUint32(root.offset + 120, newRootSize, true);
	output.view.setUint32(0x3c, output.chains[1]![0] ?? END, true);
	output.view.setUint32(0x40, newMiniFatCount, true);
	output.view.setUint32(target.offset + 116, count ? firstMini : END, true);
	output.view.setUint32(target.offset + 120, replacement.length, true);
	return output.bytes;
}
