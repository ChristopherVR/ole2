/* Adapted from ChristopherVR/pptx-viewer packages/core/src/core/utils, Apache-2.0. Original source: https://github.com/ChristopherVR/pptx-viewer. */
/**
 * OLE2 compound binary file reader.
 *
 * Parses an OLE2 container to extract named streams (e.g.
 * "EncryptionInfo", "EncryptedPackage").
 *
 * By default the reader accepts the irregularities real writers produce and
 * that cannot change stream content: a byte-order mark other than 0xFFFE,
 * trailing bytes or a short final sector, v3 stream-size high bits, a
 * mini-FAT sector count that disagrees with its chain, a chain whose
 * terminator is malformed after all advertised bytes were read, and stale
 * directory slots no storage reaches. Each is reported in `warnings`.
 * Stream bytes are never truncated, zero-filled or invented.
 * `{strict: true}` rejects all of them.
 *
 * Reference: [MS-CFB] Compound Binary File Format
 * @see https://docs.microsoft.com/en-us/openspecs/windows_protocols/ms-cfb
 *
 * @module ole2-parser-read
 */

import type { Ole2File, Ole2DirectoryEntry, Ole2ParseWarning, Ole2ReadOptions } from './ole2-parser-types.js';
import {
	OLE_MAGIC,
	ENDOFCHAIN,
	MAXREGSECT,
	ENTRY_TYPE_STREAM,
	ENTRY_TYPE_ROOT,
	Ole2ParseError,
} from './ole2-parser-types.js';
import { findDirectoryEntry, findStreamByName, readDirectory } from './ole2-directory.js';

/**
 * Parse an OLE2 compound binary file.
 *
 * @param buffer - Raw bytes of the OLE2 file.
 * @param options - Reader options; see `Ole2ReadOptions.strict`.
 * @returns Parsed OLE2 file with stream access.
 * @throws Ole2ParseError if the file is not a valid OLE2 container.
 */
export function parseOle2(buffer: ArrayBuffer | Uint8Array, options: Ole2ReadOptions = {}): Ole2File {
	const strict = options.strict === true;
	const data = buffer instanceof Uint8Array ? buffer : new Uint8Array(buffer);
	const view = new DataView(data.buffer, data.byteOffset, data.byteLength);
	const warnings: Ole2ParseWarning[] = [];
	if (data.length < 512) {
		throw new Ole2ParseError('Truncated OLE2 compound file header');
	}

	// Validate magic signature
	for (let i = 0; i < OLE_MAGIC.length; i++) {
		if (data[i] !== OLE_MAGIC[i]) {
			throw new Ole2ParseError('Not a valid OLE2 compound file');
		}
	}

	// Read header fields
	const majorVersion = view.getUint16(0x1a, true);
	const byteOrder = view.getUint16(0x1c, true);

	if (byteOrder !== 0xfffe) {
		// Some writers emit 0xFFFF; the format is little-endian regardless.
		if (strict) throw new Ole2ParseError('Invalid byte order mark');
		warnings.push({ code: 'byte-order', message: `Ignored byte order mark 0x${byteOrder.toString(16)}` });
	}

	const sectorSizePower = view.getUint16(0x1e, true);
	const miniSectorSizePower = view.getUint16(0x20, true);
	if (
		(majorVersion !== 3 && majorVersion !== 4) ||
		sectorSizePower !== (majorVersion === 3 ? 9 : 12) ||
		miniSectorSizePower !== 6
	) {
		throw new Ole2ParseError('Unsupported OLE2 version or sector size');
	}
	const sectorSize = 1 << sectorSizePower;
	const miniSectorSize = 1 << miniSectorSizePower;
	if (data.length < sectorSize) throw new Ole2ParseError('Truncated OLE2 sector');
	if (data.length % sectorSize !== 0) {
		// Trailing bytes, or a final sector some writers trim. Bytes beyond the
		// physical end are never read as zeros; see `copySector`.
		if (strict) throw new Ole2ParseError('Truncated OLE2 sector');
		warnings.push({ code: 'unaligned-length', message: `File length ${data.length} is not a multiple of ${sectorSize}` });
	}
	const sectorCount = Math.ceil(data.length / sectorSize) - 1;

	const totalFATSectors = view.getUint32(0x2c, true);
	const firstDirectorySector = view.getUint32(0x30, true);
	const miniStreamCutoff = view.getUint32(0x38, true);
	const firstMiniFATSector = view.getUint32(0x3c, true);
	const totalMiniFATSectors = view.getUint32(0x40, true);
	const firstDIFATSector = view.getUint32(0x44, true);
	const totalDIFATSectors = view.getUint32(0x48, true);
	if (totalFATSectors > sectorCount || totalMiniFATSectors > sectorCount || totalDIFATSectors > sectorCount) {
		throw new Ole2ParseError('Allocation table count exceeds file size');
	}

	function sectorOffset(sector: number): number {
		return (sector + 1) * sectorSize;
	}

	/** A whole physical sector. Allocation metadata must never be partial. */
	function readSector(sector: number): Uint8Array {
		const offset = sectorOffset(sector);
		if (offset + sectorSize > data.length) {
			throw new Ole2ParseError(
				`Sector ${sector} at offset ${offset} exceeds file size ${data.length}`,
			);
		}
		return data.subarray(offset, offset + sectorSize);
	}

	/** Copy up to `want` bytes of a sector; only the physically present ones. */
	function copySector(sector: number, target: Uint8Array, at: number, want: number): void {
		const offset = sectorOffset(sector);
		if (sector >= sectorCount || offset + want > data.length) {
			throw new Ole2ParseError(
				`Sector ${sector} at offset ${offset} exceeds file size ${data.length}`,
			);
		}
		target.set(data.subarray(offset, offset + want), at);
	}

	// Build the FAT (File Allocation Table)
	// First 109 DIFAT entries are in the header at offset 0x4C
	const fatSectors: number[] = [];
	for (let i = 0; i < 109 && fatSectors.length < totalFATSectors; i++) {
		const sector = view.getUint32(0x4c + i * 4, true);
		if (sector <= MAXREGSECT) {
			fatSectors.push(sector);
		}
	}

	// Read additional DIFAT sectors if needed
	let difatSector = firstDIFATSector;
	const difatSeen = new Set<number>();
	for (let d = 0; d < totalDIFATSectors && difatSector <= MAXREGSECT; d++) {
		if (difatSeen.has(difatSector)) throw new Ole2ParseError('Circular DIFAT chain');
		difatSeen.add(difatSector);
		const difatData = readSector(difatSector);
		const difatView = new DataView(difatData.buffer, difatData.byteOffset, difatData.byteLength);
		const entriesPerSector = (sectorSize - 4) / 4;
		for (let i = 0; i < entriesPerSector && fatSectors.length < totalFATSectors; i++) {
			const sector = difatView.getUint32(i * 4, true);
			if (sector <= MAXREGSECT) {
				fatSectors.push(sector);
			}
		}
		// Last 4 bytes of DIFAT sector point to next DIFAT sector
		difatSector = difatView.getUint32(sectorSize - 4, true);
	}
	if (difatSeen.size !== totalDIFATSectors || fatSectors.length !== totalFATSectors ||
		new Set(fatSectors).size !== fatSectors.length || fatSectors.some((id) => difatSeen.has(id))) {
		throw new Ole2ParseError('Invalid DIFAT allocation table');
	}

	// FAT padding cannot address sectors beyond the physical file. Validate
	// every listed FAT sector, but materialize only addressable entries.
	const fatEntries: number[] = [];
	for (const fatSector of fatSectors) {
		const fatData = readSector(fatSector);
		const fatView = new DataView(fatData.buffer, fatData.byteOffset, fatData.byteLength);
		for (let i = 0; i < sectorSize / 4 && fatEntries.length < sectorCount; i++) {
			fatEntries.push(fatView.getUint32(i * 4, true));
		}
	}

	/**
	 * Follow an allocation chain. With `needed`, collection stops once that
	 * many sectors are held; a malformed terminator after that point is
	 * tolerated (lenient) because every advertised byte was already found.
	 */
	function followChain(table: readonly number[], start: number, label: string, needed?: number): number[] {
		const ids: number[] = [];
		const visited = new Set<number>();
		let current = start;
		while (current <= MAXREGSECT) {
			if (!strict && needed !== undefined && ids.length >= needed) return ids;
			if (visited.has(current)) {
				throw new Ole2ParseError(`Circular reference in ${label} chain at sector ${current}`);
			}
			visited.add(current);
			ids.push(current);
			if (table[current] === undefined) throw new Ole2ParseError(`Missing ${label} chain entry`);
			current = table[current]!;
		}
		if (current !== ENDOFCHAIN) {
			if (strict || needed === undefined || ids.length < needed) {
				throw new Ole2ParseError(`Invalid ${label} chain terminator`);
			}
			warnings.push({ code: 'chain-terminator', message: `Accepted ${label} chain ending in 0x${current.toString(16)}` });
		}
		return ids;
	}

	/** Allocation metadata (directory, mini FAT): whole sectors only. */
	function readMetadataChain(startSector: number, label: string): Uint8Array {
		if (startSector > MAXREGSECT) return new Uint8Array(0);
		const ids = followChain(fatEntries, startSector, label === 'directory' ? 'FAT' : label);
		const result = new Uint8Array(ids.length * sectorSize);
		ids.forEach((id, i) => result.set(readSector(id), i * sectorSize));
		return result;
	}

	/**
	 * Read a stream, trimming to actual size.
	 */
	function readStream(startSector: number, size: number): Uint8Array {
		if (size === 0) return new Uint8Array(0);
		if (size > sectorCount * sectorSize) throw new Ole2ParseError('Stream size exceeds file size');
		const needed = Math.ceil(size / sectorSize);
		const ids = followChain(fatEntries, startSector, 'FAT', needed);
		if (ids.length < needed) throw new Ole2ParseError('Truncated stream sector chain');
		const result = new Uint8Array(size);
		for (let i = 0; i < needed; i++) {
			copySector(ids[i]!, result, i * sectorSize, Math.min(sectorSize, size - i * sectorSize));
		}
		return result;
	}

	// Read directory entries
	if (firstDirectorySector > MAXREGSECT) throw new Ole2ParseError('Missing directory chain');
	const entries: Ole2DirectoryEntry[] = readDirectory(
		readMetadataChain(firstDirectorySector, 'directory'),
		majorVersion,
		strict,
		warnings,
	);
	for (const entry of entries) {
		if (majorVersion === 4 && (!Number.isSafeInteger(entry.size) || entry.size > sectorCount * sectorSize)) {
			throw new Ole2ParseError('Stream size exceeds supported file bounds');
		}
	}

	// The root entry's stream is the mini-stream container
	const rootEntry = entries.find((e) => e.type === ENTRY_TYPE_ROOT)!;

	let miniStreamData: Uint8Array | undefined;
	if (rootEntry.startSector <= MAXREGSECT) {
		miniStreamData = readStream(rootEntry.startSector, rootEntry.size);
	}
	// Only mini sectors inside the root allocation can be addressed. Table
	// padding remains legal and never becomes an unbounded JS-number array.
	const miniFatEntries: number[] = [];
	if (firstMiniFATSector > MAXREGSECT && totalMiniFATSectors > 0) {
		if (strict) throw new Ole2ParseError('Unexpected mini FAT chain length');
		warnings.push({ code: 'mini-fat-count', message: `Header lists ${totalMiniFATSectors} mini FAT sectors but no chain` });
	}
	if (firstMiniFATSector <= MAXREGSECT && (totalMiniFATSectors > 0 || !strict)) {
		const miniFatRaw = readMetadataChain(firstMiniFATSector, 'mini FAT');
		if (miniFatRaw.length !== totalMiniFATSectors * sectorSize) {
			// Writers occasionally miscount; the chain itself is authoritative.
			if (strict) throw new Ole2ParseError('Unexpected mini FAT chain length');
			warnings.push({
				code: 'mini-fat-count',
				message: `Header lists ${totalMiniFATSectors} mini FAT sectors; chain has ${miniFatRaw.length / sectorSize}`,
			});
		}
		const miniFatView = new DataView(miniFatRaw.buffer, miniFatRaw.byteOffset, miniFatRaw.byteLength);
		const miniCapacity = Math.floor((miniStreamData?.length ?? 0) / miniSectorSize);
		for (let i = 0; i < miniFatRaw.length / 4 && i < miniCapacity; i++) {
			miniFatEntries.push(miniFatView.getUint32(i * 4, true));
		}
	}

	/**
	 * Read a mini-stream, following the mini FAT chain.
	 */
	function readMiniStream(startSector: number, size: number): Uint8Array {
		if (size === 0) return new Uint8Array(0);
		if (!miniStreamData) {
			throw new Ole2ParseError('Mini stream container not found');
		}
		const needed = Math.ceil(size / miniSectorSize);
		const capacity = miniStreamData.length / miniSectorSize;
		let ids: number[];
		try {
			ids = followChain(miniFatEntries, startSector, 'mini FAT', needed);
		} catch (error) {
			if (error instanceof Ole2ParseError && /Missing mini FAT/.test(error.message)) {
				throw new Ole2ParseError('Mini stream sector exceeds allocation bounds');
			}
			throw error;
		}
		const result = new Uint8Array(size);
		for (let i = 0; i < ids.length && i < needed; i++) {
			const id = ids[i]!;
			if (id >= capacity || (id + 1) * miniSectorSize > miniStreamData.length) {
				throw new Ole2ParseError('Mini stream sector exceeds allocation bounds');
			}
			const want = Math.min(miniSectorSize, size - i * miniSectorSize);
			result.set(miniStreamData.subarray(id * miniSectorSize, id * miniSectorSize + want), i * miniSectorSize);
		}
		if (ids.length < needed) throw new Ole2ParseError('Truncated mini stream sector chain');
		return result;
	}

	function readEntry(entry: Ole2DirectoryEntry): Uint8Array {
		if (entry.size < miniStreamCutoff && entry.type !== ENTRY_TYPE_ROOT) {
			return readMiniStream(entry.startSector, entry.size);
		}
		return readStream(entry.startSector, entry.size);
	}
	const readable = (entry: Ole2DirectoryEntry) => entry.type === ENTRY_TYPE_STREAM || entry.type === ENTRY_TYPE_ROOT;

	return {
		entries,
		warnings,
		getStream(name: string): Uint8Array | undefined {
			const entry = findStreamByName(entries, name, readable);
			return entry ? readEntry(entry) : undefined;
		},
		getStreamByPath(path: string | readonly string[]): Uint8Array | undefined {
			const entry = findDirectoryEntry(entries, path);
			return entry && entry.type === ENTRY_TYPE_STREAM ? readEntry(entry) : undefined;
		},
		getStreamById(id: number): Uint8Array | undefined {
			const entry = entries.find((candidate) => candidate.id === id);
			return entry && readable(entry) ? readEntry(entry) : undefined;
		},
		findEntry(path: string | readonly string[]): Ole2DirectoryEntry | undefined {
			return findDirectoryEntry(entries, path);
		},
	};
}
