/** Append allocations without relocating any existing physical sector. */
export function appendCompoundAllocations(
	input: Uint8Array,
	existingFatIds: readonly number[],
	payloads: readonly Uint8Array[],
	sectorSize = 512,
): {
	bytes: Uint8Array;
	view: DataView;
	chains: number[][];
	writeFat: (id: number, value: number) => void;
} | undefined {
	const sectorCount = input.length / sectorSize - 1;
	const entriesPerFat = sectorSize / 4;
	const counts = payloads.map((payload) => Math.ceil(payload.length / sectorSize));
	const payloadSectors = counts.reduce((sum, count) => sum + count, 0);
	let fatCount = existingFatIds.length;
	while (sectorCount + payloadSectors + fatCount - existingFatIds.length > fatCount * entriesPerFat) fatCount++;
	if (fatCount > 109) return undefined;
	const totalSectors = sectorCount + payloadSectors + fatCount - existingFatIds.length;
	const bytes = new Uint8Array((totalSectors + 1) * sectorSize);
	bytes.set(input);
	const view = new DataView(bytes.buffer), fatIds = [...existingFatIds];
	for (let i = existingFatIds.length; i < fatCount; i++) {
		const id = sectorCount + payloadSectors + i - existingFatIds.length;
		fatIds.push(id);
		bytes.fill(0xff, (id + 1) * sectorSize, (id + 2) * sectorSize);
		view.setUint32(0x4c + i * 4, id, true);
	}
	view.setUint32(0x2c, fatCount, true);
	const writeFat = (id: number, value: number) => view.setUint32((fatIds[Math.floor(id / entriesPerFat)]! + 1) * sectorSize + (id % entriesPerFat) * 4, value, true);
	const chains: number[][] = [];
	let next = sectorCount;
	for (let payload = 0; payload < payloads.length; payload++) {
		bytes.set(payloads[payload]!, (next + 1) * sectorSize);
		const chain = Array.from({ length: counts[payload]! }, (_, i) => next + i);
		chain.forEach((id, i) => writeFat(id, chain[i + 1] ?? 0xfffffffe));
		chains.push(chain);
		next += counts[payload]!;
	}
	for (const id of fatIds) writeFat(id, 0xfffffffd);
	return { bytes, view, chains, writeFat };
}

/** Read a bounded chain; table entries beyond physical capacity are absent. */
export function followCompoundAllocation(table: readonly number[], start: number): number[] {
	const ids: number[] = [], seen = new Set<number>();
	for (let id = start; id !== 0xfffffffe; id = table[id]!) {
		if (id >= table.length || seen.has(id)) throw new Error('Invalid allocation chain');
		seen.add(id);
		ids.push(id);
	}
	return ids;
}
