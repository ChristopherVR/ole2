/** Safe, same-length mutation of a single stream in an MS-CFB container. */
const END = 0xfffffffe;
const MAX = 0xfffffffa;
const NIL = 0xffffffff;
const SIZE = 128;

type Entry = {
	id: number;
	name: string;
	type: number;
	left: number;
	right: number;
	child: number;
	start: number;
	size: number;
};

function chain(table: readonly number[], start: number, limit: number): number[] {
	if (start === END && limit === 0) return [];
	const out: number[] = [],
		seen = new Set<number>();
	let id = start;
	while (id !== END) {
		if (id > MAX || id >= table.length || seen.has(id) || out.length >= limit)
			throw new Error('Invalid CFB chain');
		seen.add(id);
		out.push(id);
		id = table[id]!;
	}
	if (out.length !== limit) throw new Error('Unexpected CFB chain length');
	return out;
}

function chainToEnd(table: readonly number[], start: number, max: number): number[] {
	if (start === END) return [];
	const out: number[] = [],
		seen = new Set<number>();
	let id = start;
	while (id !== END) {
		if (id > MAX || id >= table.length || seen.has(id) || out.length >= max)
			throw new Error('Invalid CFB chain');
		seen.add(id);
		out.push(id);
		id = table[id]!;
	}
	return out;
}

/** Leading `needed` ids of a chain; a long or badly terminated tail past
 * the advertised size is never followed. */
function chainPrefix(table: readonly number[], start: number, needed: number): number[] {
	const out: number[] = [],
		seen = new Set<number>();
	let id = start;
	while (out.length < needed) {
		if (id > MAX || id >= table.length || seen.has(id)) throw new Error('Invalid CFB chain');
		seen.add(id);
		out.push(id);
		id = table[id]!;
	}
	return out;
}

/**
 * Read-only path: validate just the target's allocation. Unrelated damage
 * elsewhere in the container cannot hide a healthy stream, and bytes past the
 * physical end of a trimmed file are refused rather than zero-filled.
 */
function readTarget(
	input: Uint8Array,
	target: Entry,
	root: Entry,
	fat: readonly number[],
	sector: (id: number) => Uint8Array,
	sectorCount: number,
	sectorSize: number,
	miniSize: number,
	read32: (offset: number) => number,
): Uint8Array {
	const size = target.size;
	if (size === 0) return new Uint8Array(0);
	const out = new Uint8Array(size);
	const copyPhysical = (at: number, want: number, into: number) => {
		if (at + want > input.length) throw new Error('Stream data exceeds file size');
		out.set(input.subarray(at, at + want), into);
	};
	if (size >= read32(0x38)) {
		if (size > sectorCount * sectorSize) throw new Error('Stream size exceeds file size');
		chainPrefix(fat, target.start, Math.ceil(size / sectorSize)).forEach((id, i) => {
			if (id >= sectorCount) throw new Error('Invalid sector');
			copyPhysical((id + 1) * sectorSize, Math.min(sectorSize, size - i * sectorSize), i * sectorSize);
		});
		return out;
	}
	// The mini-FAT chain is authoritative; its header count is advisory here.
	const miniFatSectors = chainToEnd(fat, read32(0x3c), sectorCount);
	const miniCapacity = Math.floor(Math.min(root.size, sectorCount * sectorSize) / miniSize);
	const miniFat: number[] = [];
	for (const id of miniFatSectors) {
		const bytes = sector(id),
			dv = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
		for (let j = 0; j < sectorSize / 4 && miniFat.length < miniCapacity; j++) miniFat.push(dv.getUint32(j * 4, true));
	}
	const rootSectors = chainPrefix(fat, root.start, Math.ceil(root.size / sectorSize));
	chainPrefix(miniFat, target.start, Math.ceil(size / miniSize)).forEach((id, i) => {
		const start = id * miniSize,
			want = Math.min(miniSize, size - i * miniSize);
		if (start + want > root.size) throw new Error('Mini stream out of bounds');
		const rootSector = rootSectors[Math.floor(start / sectorSize)]!;
		if (rootSector >= sectorCount) throw new Error('Invalid sector');
		copyPhysical((rootSector + 1) * sectorSize + (start % sectorSize), want, i * miniSize);
	});
	return out;
}

/**
 * Replace a stream payload without reallocating sectors. The path is relative
 * to the root storage and includes every storage name followed by the stream.
 * Returns the original input reference when validation fails.
 */
function processCompoundFileStream(
	input: Uint8Array,
	path: readonly string[],
	replacement?: Uint8Array,
): Uint8Array | undefined {
	const reject = () => (replacement ? input : undefined);
	try {
		if (path.length === 0 || path.some((part) => !part) || input.length < 512) return reject();
		const view = new DataView(input.buffer, input.byteOffset, input.byteLength);
		const read32 = (offset: number) => view.getUint32(offset, true);
		if (
			input[0] !== 0xd0 ||
			input[1] !== 0xcf ||
			input[2] !== 0x11 ||
			input[3] !== 0xe0 ||
			input[4] !== 0xa1 ||
			input[5] !== 0xb1 ||
			input[6] !== 0x1a ||
			input[7] !== 0xe1 ||
			// Readers tolerate a non-canonical byte-order mark; edits do not.
			(replacement && view.getUint16(0x1c, true) !== 0xfffe)
		)
			return reject();
		const major = view.getUint16(0x1a, true),
			shift = view.getUint16(0x1e, true),
			miniShift = view.getUint16(0x20, true);
		if ((major !== 3 && major !== 4) || shift !== (major === 3 ? 9 : 12) || miniShift !== 6)
			return reject();
		const sectorSize = 1 << shift,
			miniSize = 1 << miniShift;
		// Reads accept trailing bytes or a trimmed final sector, but never
		// return bytes beyond the physical end; edits need a canonical file.
		if (input.length < sectorSize || (replacement && input.length % sectorSize !== 0)) return reject();
		const sectorCount = Math.ceil(input.length / sectorSize) - 1;
		const offset = (id: number) => (id + 1) * sectorSize;
		const sector = (id: number) => {
			if (id > MAX || id >= sectorCount || offset(id) + sectorSize > input.length) throw new Error('Invalid sector');
			return input.subarray(offset(id), offset(id) + sectorSize);
		};
		const fatCount = read32(0x2c),
			difatCount = read32(0x48),
			difatStart = read32(0x44);
		const fatIds: number[] = [];
		for (let i = 0; i < 109 && fatIds.length < fatCount; i++) {
			const id = read32(0x4c + 4 * i);
			if (id <= MAX) fatIds.push(id);
		}
		let difat = difatStart;
		const difatIds = new Set<number>();
		for (let i = 0; i < difatCount; i++) {
			if (difatIds.has(difat) || difat > MAX) return reject();
			difatIds.add(difat);
			const bytes = sector(difat),
				dv = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
			for (let j = 0; j < (sectorSize - 4) / 4 && fatIds.length < fatCount; j++) {
				const id = dv.getUint32(j * 4, true);
				if (id <= MAX) fatIds.push(id);
			}
			difat = dv.getUint32(sectorSize - 4, true);
		}
		if (
			fatIds.length !== fatCount ||
			new Set(fatIds).size !== fatIds.length ||
			fatIds.some((id) => difatIds.has(id))
		)
			return reject();
		const fat: number[] = [];
		for (const id of fatIds) {
			const bytes = sector(id),
				dv = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
			for (let j = 0; j < sectorSize / 4 && fat.length < sectorCount; j++) fat.push(dv.getUint32(j * 4, true));
		}
		const exactChain = (table: readonly number[], start: number, bytes: number, unit: number) =>
			chain(table, start, Math.ceil(bytes / unit));
		const dirStart = read32(0x30),
			dirSectors = chainToEnd(fat, dirStart, sectorCount);
		const dirBytes = new Uint8Array(dirSectors.length * sectorSize);
		dirSectors.forEach((id, i) => dirBytes.set(sector(id), i * sectorSize));
		const dvDir = new DataView(dirBytes.buffer);
		const entries: Entry[] = [];
		for (let p = 0, id = 0; p + SIZE <= dirBytes.length; p += SIZE, id++) {
			const type = dirBytes[p + 66]!;
			if (type === 0) continue;
			const nameLength = dvDir.getUint16(p + 64, true);
			if (nameLength < 2 || nameLength > 64 || nameLength % 2) return reject();
			let name = '';
			for (let i = 0; i < nameLength - 2; i += 2)
				name += String.fromCharCode(dvDir.getUint16(p + i, true));
			const sizeLow = dvDir.getUint32(p + 120, true),
				sizeHigh = dvDir.getUint32(p + 124, true);
			// 64-bit stream sizes are unsupported. v3 readers ignore the high
			// bits, which old writers left uninitialized ([MS-CFB] 2.6.3).
			if (sizeHigh !== 0 && (replacement || major === 4)) return reject();
			entries.push({
				id,
				name,
				type,
				left: dvDir.getUint32(p + 68, true),
				right: dvDir.getUint32(p + 72, true),
				child: dvDir.getUint32(p + 76, true),
				start: dvDir.getUint32(p + 116, true),
				size: sizeLow,
			});
		}
		const byId = new Map(entries.map((e) => [e.id, e]));
		const root = entries.find((e) => e.type === 5);
		if (!root) return reject();
		let parent = root,
			target: Entry | undefined;
		for (let partIndex = 0; partIndex < path.length; partIndex++) {
			const matches: Entry[] = [],
				seen = new Set<number>();
			const visit = (id: number) => {
				if (id === NIL) return;
				if (seen.has(id)) throw new Error('Cyclic directory tree');
				seen.add(id);
				const e = byId.get(id);
				if (!e) throw new Error('Invalid directory ID');
				visit(e.left);
				if (e.name.toLocaleLowerCase() === path[partIndex]!.toLocaleLowerCase()) matches.push(e);
				visit(e.right);
			};
			visit(parent.child);
			if (matches.length !== 1) return reject();
			target = matches[0];
			if (partIndex < path.length - 1) {
				if (target.type !== 1) return reject();
				parent = target;
			} else if (target.type !== 2 || (replacement && replacement.length !== target.size))
				return reject();
		}
		if (!target) return reject();
		if (!replacement) return readTarget(input, target, root, fat, sector, sectorCount, sectorSize, miniSize, read32);
		const names = entries.map((e) => e.name.toLowerCase());
		if (replacement && (names.includes('encryptioninfo') || names.includes('encryptedpackage')))
			return reject();
		for (const name of replacement ? ['book', 'workbook'] : []) {
			if (!names.includes(name)) continue;
			const workbook = readCompoundFileStream(input, [name]);
			if (!workbook) continue;
			for (let p = 0; p + 4 <= workbook.length;) {
				const record = workbook[p]! | (workbook[p + 1]! << 8);
				const length = workbook[p + 2]! | (workbook[p + 3]! << 8);
				if (record === 0x002f) return reject(); // BIFF FILEPASS
				if (p + 4 + length > workbook.length || record === 0x000a) break;
				p += 4 + length;
			}
		}
		if (replacement && names.includes('worddocument')) {
			const word = readCompoundFileStream(input, ['WordDocument']);
			if (word && word.length >= 12 && ((word[10]! | (word[11]! << 8)) & 0x0100) !== 0)
				return reject();
		}
		const miniCutoff = read32(0x38),
			miniFatCount = read32(0x40),
			miniFatStart = read32(0x3c);
		const miniFatSectors = miniFatCount ? chain(fat, miniFatStart, miniFatCount) : [];
		const metadataSectors = [...fatIds, ...difatIds, ...dirSectors, ...miniFatSectors];
		if (new Set(metadataSectors).size !== metadataSectors.length) return reject();
		const miniFat: number[] = [];
		const miniCapacity = Math.floor(Math.min(root.size, sectorCount * sectorSize) / miniSize);
		for (const id of miniFatSectors) {
			const bytes = sector(id),
				dv = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
			for (let j = 0; j < sectorSize / 4 && miniFat.length < miniCapacity; j++) miniFat.push(dv.getUint32(j * 4, true));
		}
		const reserved = new Set<number>([...fatIds, ...difatIds, ...dirSectors, ...miniFatSectors]);
		const regularUsed = new Set<number>(reserved);
		const rootMini = entries.find((e) => e.type === 5)!;
		for (const entry of entries) {
			if ((entry.type !== 2 && entry.type !== 5) || entry.size === 0) continue;
			if (entry.type === 2 && entry.size < miniCutoff) continue;
			for (const id of exactChain(fat, entry.start, entry.size, sectorSize)) {
				sector(id);
				if (regularUsed.has(id)) return reject();
				regularUsed.add(id);
			}
		}
		const miniUsed = new Set<number>();
		for (const entry of entries) {
			if (entry.type !== 2 || entry.size === 0 || entry.size >= miniCutoff) continue;
			const ids = exactChain(miniFat, entry.start, entry.size, miniSize);
			for (const id of ids) {
				if ((id + 1) * miniSize > rootMini.size || miniUsed.has(id)) return reject();
				miniUsed.add(id);
			}
		}
		if (!replacement) {
			if (target.size >= miniCutoff) {
				const ids = exactChain(fat, target.start, target.size, sectorSize);
				const result = new Uint8Array(ids.length * sectorSize);
				ids.forEach((id, i) => result.set(sector(id), i * sectorSize));
				return result.slice(0, target.size);
			}
			const ids = exactChain(miniFat, target.start, target.size, miniSize);
			const rootSectors = exactChain(fat, rootMini.start, rootMini.size, sectorSize);
			const result = new Uint8Array(ids.length * miniSize);
			ids.forEach((id, i) => {
				const start = id * miniSize;
				if (start + miniSize > rootMini.size) throw new Error('Mini stream out of bounds');
				const rootSector = rootSectors[Math.floor(start / sectorSize)]!;
				const source = sector(rootSector).subarray(
					start % sectorSize,
					(start % sectorSize) + miniSize,
				);
				result.set(source, i * miniSize);
			});
			return result.slice(0, target.size);
		}
		// A validated empty stream has no payload allocation to overwrite.
		if (replacement.length === 0) return input.slice();
		const output = input.slice();
		if (target.size >= miniCutoff) {
			const sectors = exactChain(fat, target.start, target.size, sectorSize);
			sectors.forEach((id) => sector(id));
			for (let i = 0; i < sectors.length; i++)
				output.set(
					replacement.subarray(i * sectorSize, Math.min((i + 1) * sectorSize, replacement.length)),
					offset(sectors[i]!),
				);
		} else {
			if (!miniFat.length) return reject();
			const miniSectors = exactChain(miniFat, target.start, target.size, miniSize);
			const rootMini = entries.find((e) => e.type === 5)!;
			const rootSectors = exactChain(fat, rootMini.start, rootMini.size, sectorSize);
			rootSectors.forEach((id) => sector(id));
			const rootData = new Uint8Array(rootMini.size);
			rootSectors.forEach((id, i) =>
				rootData.set(
					sector(id).subarray(0, Math.min(sectorSize, rootData.length - i * sectorSize)),
					i * sectorSize,
				),
			);
			for (let i = 0; i < replacement.length; i++) {
				const miniOffset = miniSectors[Math.floor(i / miniSize)]! * miniSize + (i % miniSize);
				if (miniOffset >= rootData.length) return reject();
				const rootSectorIndex = Math.floor(miniOffset / sectorSize);
				const rootSector = rootSectors[rootSectorIndex];
				if (rootSector === undefined) return reject();
				const physical = offset(rootSector) + (miniOffset % sectorSize);
				output[physical] = replacement[i]!;
			}
		}
		return output;
	} catch {
		return reject();
	}
}

/** Read a stream by its storage path; undefined means missing or malformed. */
export function readCompoundFileStream(
	input: Uint8Array,
	path: readonly string[],
): Uint8Array | undefined {
	return processCompoundFileStream(input, path);
}

/** Replace a stream payload without reallocating sectors. The path is relative
 * to the root storage and includes every storage name followed by the stream.
 * Returns the original input reference when validation fails. */
export function replaceCompoundFileStream(
	input: Uint8Array,
	path: readonly string[],
	replacement: Uint8Array,
): Uint8Array {
	return processCompoundFileStream(input, path, replacement) ?? input;
}
