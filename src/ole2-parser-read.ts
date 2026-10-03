/* Adapted from ChristopherVR/pptx-viewer packages/core/src/core/utils, Apache-2.0. Original source: https://github.com/ChristopherVR/pptx-viewer. */
/**
 * OLE2 compound binary file reader.
 *
 * Parses an OLE2 container to extract named streams (e.g.
 * "EncryptionInfo", "EncryptedPackage").
 *
 * Reference: [MS-CFB] Compound Binary File Format
 * @see https://docs.microsoft.com/en-us/openspecs/windows_protocols/ms-cfb
 *
 * @module ole2-parser-read
 */

import type { Ole2File, Ole2DirectoryEntry } from './ole2-parser-types.js';
import {
	OLE_MAGIC,
	ENDOFCHAIN,
	MAXREGSECT,
	ENTRY_TYPE_EMPTY,
	ENTRY_TYPE_STREAM,
	ENTRY_TYPE_ROOT,
	DIR_ENTRY_SIZE,
	Ole2ParseError,
} from './ole2-parser-types.js';

/**
 * Parse an OLE2 compound binary file from an ArrayBuffer.
 *
 * @param buffer - Raw bytes of the OLE2 file.
 * @returns Parsed OLE2 file with stream access.
 * @throws Ole2ParseError if the file is not a valid OLE2 container.
 */
export function parseOle2(buffer: ArrayBuffer): Ole2File {
	const data = new Uint8Array(buffer);
	const view = new DataView(buffer);
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
	const _minorVersion = view.getUint16(0x18, true);
	const majorVersion = view.getUint16(0x1a, true);
	const byteOrder = view.getUint16(0x1c, true);

	if (byteOrder !== 0xfffe) {
		throw new Ole2ParseError('Invalid byte order mark');
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
	if (data.length < sectorSize || data.length % sectorSize !== 0) {
		throw new Ole2ParseError('Truncated OLE2 sector');
	}
	const sectorCount = data.length / sectorSize - 1;

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

	// Helper: convert sector index to file offset
	function sectorOffset(sector: number): number {
		return (sector + 1) * sectorSize;
	}

	// Read sector data
	function readSector(sector: number): Uint8Array {
		const offset = sectorOffset(sector);
		if (offset + sectorSize > data.length) {
			throw new Ole2ParseError(
				`Sector ${sector} at offset ${offset} exceeds file size ${data.length}`,
			);
		}
		return data.subarray(offset, offset + sectorSize);
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
	 * Read a chain of sectors following the FAT.
	 */
	function readSectorChain(startSector: number): Uint8Array {
		const sectors: Uint8Array[] = [];
		let current = startSector;
		const visited = new Set<number>();

		while (current <= MAXREGSECT) {
			if (visited.has(current)) {
				throw new Ole2ParseError(`Circular reference in FAT chain at sector ${current}`);
			}
			visited.add(current);
			sectors.push(readSector(current));
			if (fatEntries[current] === undefined) throw new Ole2ParseError('Missing FAT chain entry');
			current = fatEntries[current]!;
		}
		if (current !== ENDOFCHAIN) throw new Ole2ParseError('Invalid FAT chain terminator');

		// Concatenate all sectors
		const totalLength = sectors.length * sectorSize;
		const result = new Uint8Array(totalLength);
		let offset = 0;
		for (const sector of sectors) {
			result.set(sector, offset);
			offset += sectorSize;
		}
		return result;
	}

	/**
	 * Read a stream, trimming to actual size.
	 */
	function readStream(startSector: number, size: number): Uint8Array {
		if (size === 0) return new Uint8Array(0);
		if (size > sectorCount * sectorSize) throw new Ole2ParseError('Stream size exceeds file size');
		const raw = readSectorChain(startSector);
		if (raw.length < size) throw new Ole2ParseError('Truncated stream sector chain');
		return raw.subarray(0, size);
	}

	// Read directory entries
	const dirRaw = readSectorChain(firstDirectorySector);
	const numEntries = Math.floor(dirRaw.length / DIR_ENTRY_SIZE);
	const entries: Ole2DirectoryEntry[] = [];

	for (let i = 0; i < numEntries; i++) {
		const entryOffset = i * DIR_ENTRY_SIZE;
		const entryView = new DataView(dirRaw.buffer, dirRaw.byteOffset + entryOffset, DIR_ENTRY_SIZE);

		const nameLen = entryView.getUint16(64, true);
		const objectType = entryView.getUint8(66);

		if (objectType === ENTRY_TYPE_EMPTY) {
			continue;
		}
		if (nameLen < 2 || nameLen > 64 || nameLen % 2 !== 0 || entryView.getUint16(nameLen - 2, true) !== 0) {
			throw new Ole2ParseError('Invalid directory entry name');
		}

		// Name is a UTF-16LE string, nameLen includes the null terminator (in bytes)
		const nameBytes = Math.max(0, nameLen - 2);
		let name = '';
		for (let j = 0; j < nameBytes; j += 2) {
			name += String.fromCharCode(entryView.getUint16(j, true));
		}

		const leftSiblingId = entryView.getUint32(68, true);
		const rightSiblingId = entryView.getUint32(72, true);
		const childId = entryView.getUint32(76, true);
		const clsid = new Uint8Array(
			entryView.buffer.slice(entryView.byteOffset + 80, entryView.byteOffset + 96),
		);
		const startSector = entryView.getUint32(116, true);
		const sizeLow = entryView.getUint32(120, true);

		// For v4 files, size can be 64-bit
		let size = sizeLow;
		if (majorVersion === 4) {
			const sizeHigh = entryView.getUint32(124, true);
			size = sizeLow + sizeHigh * 0x100000000;
			if (!Number.isSafeInteger(size) || size > sectorCount * sectorSize) {
				throw new Ole2ParseError('Stream size exceeds supported file bounds');
			}
		}

		entries.push({
			name,
			type: objectType,
			startSector,
			size,
			childId: childId === 0xffffffff ? -1 : childId,
			leftSiblingId: leftSiblingId === 0xffffffff ? -1 : leftSiblingId,
			rightSiblingId: rightSiblingId === 0xffffffff ? -1 : rightSiblingId,
			clsid,
		});
	}

	// The root entry's stream is the mini-stream container
	const rootEntry = entries.find((e) => e.type === ENTRY_TYPE_ROOT);
	if (!rootEntry || entries.filter((e) => e.type === ENTRY_TYPE_ROOT).length !== 1) {
		throw new Ole2ParseError('Invalid root directory entry');
	}

	let miniStreamData: Uint8Array | undefined;
	if (rootEntry && rootEntry.startSector <= MAXREGSECT) {
		miniStreamData = readStream(rootEntry.startSector, rootEntry.size);
	}
	// Only mini sectors inside the root allocation can be addressed. Table
	// padding remains legal and never becomes an unbounded JS-number array.
	const miniFatEntries: number[] = [];
	if (firstMiniFATSector <= MAXREGSECT && totalMiniFATSectors > 0) {
		const miniFatRaw = readSectorChain(firstMiniFATSector);
		if (miniFatRaw.length !== totalMiniFATSectors * sectorSize) {
			throw new Ole2ParseError('Unexpected mini FAT chain length');
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

		const sectors: Uint8Array[] = [];
		let current = startSector;
		const visited = new Set<number>();

		while (current <= MAXREGSECT) {
			if (visited.has(current)) {
				throw new Ole2ParseError(`Circular reference in mini FAT chain at sector ${current}`);
			}
			visited.add(current);
			const offset = current * miniSectorSize;
			if (offset + miniSectorSize > miniStreamData.length || miniFatEntries[current] === undefined) {
				throw new Ole2ParseError('Mini stream sector exceeds allocation bounds');
			}
			sectors.push(miniStreamData.subarray(offset, offset + miniSectorSize));
			current = miniFatEntries[current]!;
		}
		if (current !== ENDOFCHAIN) throw new Ole2ParseError('Invalid mini FAT chain terminator');

		const totalLength = sectors.length * miniSectorSize;
		if (totalLength < size) throw new Ole2ParseError('Truncated mini stream sector chain');
		const result = new Uint8Array(totalLength);
		let offset = 0;
		for (const sector of sectors) {
			result.set(sector, offset);
			offset += miniSectorSize;
		}
		return result.subarray(0, size);
	}

	/**
	 * Get a named stream from the OLE2 file.
	 */
	function getStream(name: string): Uint8Array | undefined {
		const entry = entries.find(
			(e) => (e.type === ENTRY_TYPE_STREAM || e.type === ENTRY_TYPE_ROOT) && e.name === name,
		);
		if (!entry) {
			return undefined;
		}

		if (entry.size < miniStreamCutoff && entry.type !== ENTRY_TYPE_ROOT) {
			return readMiniStream(entry.startSector, entry.size);
		}
		return readStream(entry.startSector, entry.size);
	}

	return { entries, getStream };
}
