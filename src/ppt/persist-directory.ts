/**
 * UserEditAtom chain + PersistDirectoryAtom handling ([MS-PPT] 2.3.3-2.3.6).
 *
 * A .ppt file supports incremental saves: each save appends a user edit
 * (records + a PersistDirectoryAtom + a UserEditAtom). The live persist
 * object directory is built by walking the UserEditAtom chain from the most
 * recent edit backwards, then applying the persist directories from oldest
 * to newest so newer entries override older ones.
 *
 * @module ppt/persist-directory
 */

import { PptParseError, readRecordOrThrow } from './record-stream.js';
import type { PptRecord } from './record-stream.js';
import { RT } from './record-types.js';

/** Resource limits for untrusted save histories; callers may lower or raise them. */
export interface PersistDirectoryLimits {
	maxEdits?: number;
	maxEntries?: number;
}

const DEFAULT_MAX_EDITS = 10000;
const DEFAULT_MAX_ENTRIES = 1000000;

function checkedLimit(value: number): number {
	if (!Number.isSafeInteger(value) || value < 1) throw new PptParseError('Invalid persist directory resource limit');
	return value;
}

function boundedAtom(view: DataView, offset: number, type: number): PptRecord {
	const rec = readRecordOrThrow(view, offset);
	if (rec.recType !== type || rec.recVer !== 0 || rec.recInstance !== 0) {
		throw new PptParseError(`Invalid persist atom header at offset ${offset}`);
	}
	if (rec.recLen > view.byteLength - rec.dataOffset) {
		throw new PptParseError(`Persist atom payload out of bounds at offset ${offset}`);
	}
	return rec;
}

/** Parsed UserEditAtom. */
export interface UserEditAtom {
	/** Offset of the previous UserEditAtom, 0 when none. */
	offsetLastEdit: number;
	/** Offset of the PersistDirectoryAtom for this edit. */
	offsetPersistDirectory: number;
	/** Persist id of the DocumentContainer (must be 1). */
	docPersistIdRef: number;
	/** Persist id of the CryptSession10Container when the file is encrypted. */
	encryptSessionPersistIdRef: number | undefined;
}

/** Persist object directory: persist id -> stream offset. */
export type PersistDirectory = Map<number, number>;

/** Result of walking the user edit chain. */
export interface UserEditChain {
	/** Most recent UserEditAtom (the live one). */
	currentEdit: UserEditAtom;
	/** Live persist object directory. */
	directory: PersistDirectory;
}

/** Parse a UserEditAtom record at the given offset. */
export function parseUserEditAtom(view: DataView, offset: number): UserEditAtom {
	const rec = boundedAtom(view, offset, RT.UserEditAtom);
	if (rec.recLen !== 0x1c && rec.recLen !== 0x20) {
		throw new PptParseError(`UserEditAtom has invalid length ${rec.recLen}`);
	}
	const d = rec.dataOffset;
	const edit: UserEditAtom = {
		offsetLastEdit: view.getUint32(d + 8, true),
		offsetPersistDirectory: view.getUint32(d + 12, true),
		docPersistIdRef: view.getUint32(d + 16, true),
		encryptSessionPersistIdRef: rec.recLen === 0x20 ? view.getUint32(d + 28, true) : undefined,
	};
	if (view.getUint8(d + 6) !== 0 || view.getUint8(d + 7) !== 3) {
		throw new PptParseError('Unsupported UserEditAtom storage version');
	}
	if (edit.offsetLastEdit >= offset || edit.offsetPersistDirectory <= edit.offsetLastEdit || edit.offsetPersistDirectory >= offset) {
		throw new PptParseError('Invalid UserEditAtom save history offsets');
	}
	if (edit.docPersistIdRef !== 1) throw new PptParseError('Invalid document persist identifier');
	return edit;
}

/**
 * Parse a PersistDirectoryAtom record into (persistId, offset) pairs.
 *
 * Each PersistDirectoryEntry is a packed UInt32 (persistId in the low
 * 20 bits, cPersist in the high 12 bits) followed by cPersist UInt32
 * stream offsets for persistId, persistId+1, ...
 */
export function parsePersistDirectoryAtom(view: DataView, offset: number, maxEntries = DEFAULT_MAX_ENTRIES): Array<[number, number]> {
	checkedLimit(maxEntries);
	const rec = boundedAtom(view, offset, RT.PersistDirectoryAtom);
	const pairs: Array<[number, number]> = [];
	let pos = rec.dataOffset;
	const end = rec.dataOffset + rec.recLen;
	const seen = new Set<number>();
	while (pos < end) {
		if (end - pos < 4) throw new PptParseError('Truncated persist directory entry');
		const packed = view.getUint32(pos, true);
		const persistId = packed & 0xfffff;
		const cPersist = (packed >>> 20) & 0xfff;
		pos += 4;
		if (persistId > 0xffffe || cPersist === 0 || persistId + cPersist - 1 > 0xfffff) {
			throw new PptParseError('Invalid persist identifier run');
		}
		if (cPersist * 4 > end - pos) throw new PptParseError('Truncated persist offset run');
		if (pairs.length + cPersist > maxEntries) throw new PptParseError('Persist directory entry limit exceeded');
		for (let i = 0; i < cPersist; i++, pos += 4) {
			const id = persistId + i;
			if (seen.has(id)) throw new PptParseError('Duplicate persist identifier');
			seen.add(id);
			const objectOffset = view.getUint32(pos, true);
			if (objectOffset >= offset || objectOffset + 8 > view.byteLength) throw new PptParseError('Persist object offset out of bounds');
			pairs.push([id, objectOffset]);
		}
	}
	return pairs;
}

/**
 * Walk the UserEditAtom chain starting from `offsetToCurrentEdit` and build
 * the live persist object directory (newest entries win).
 *
 * @param view - DataView over the PowerPoint Document stream.
 * @param offsetToCurrentEdit - From the CurrentUserAtom.
 */
export function buildPersistDirectory(view: DataView, offsetToCurrentEdit: number, limits: PersistDirectoryLimits = {}): UserEditChain {
	const maxEdits = checkedLimit(limits.maxEdits ?? DEFAULT_MAX_EDITS);
	const maxEntries = checkedLimit(limits.maxEntries ?? DEFAULT_MAX_ENTRIES);
	const edits: UserEditAtom[] = [];
	const editOffsets: number[] = [];
	const seen = new Set<number>();
	let offset = offsetToCurrentEdit;

	while (offset !== 0) {
		if (edits.length >= maxEdits) throw new PptParseError('User edit chain limit exceeded');
		if (seen.has(offset)) {
			throw new PptParseError(`Circular UserEditAtom chain at offset ${offset}`);
		}
		seen.add(offset);
		const edit = parseUserEditAtom(view, offset);
		edits.push(edit);
		editOffsets.push(offset);
		offset = edit.offsetLastEdit;
	}

	if (edits.length === 0) {
		throw new PptParseError('No UserEditAtom found');
	}

	// Apply persist directories oldest-first so newer entries override.
	const directory: PersistDirectory = new Map();
	let entryCount = 0;
	for (let i = edits.length - 1; i >= 0; i--) {
		const edit = edits[i];
		const dirAtom = boundedAtom(view, edit.offsetPersistDirectory, RT.PersistDirectoryAtom);
		// A directory cannot consume bytes from the UserEditAtom that follows it.
		const editOffset = editOffsets[i];
		if (dirAtom.dataOffset + dirAtom.recLen > editOffset) throw new PptParseError('Persist directory overlaps user edit');
		const entries = parsePersistDirectoryAtom(view, edit.offsetPersistDirectory, maxEntries);
		entryCount += entries.length;
		if (entryCount > maxEntries) throw new PptParseError('Save history persist entry limit exceeded');
		for (const [id, off] of entries) {
			if (off < edit.offsetLastEdit) throw new PptParseError('Persist object precedes corresponding user edit');
			directory.set(id, off);
		}
	}

	return { currentEdit: edits[0], directory };
}

/**
 * Resolve a persist id to the record at its stream offset.
 */
export function readPersistRecord(
	view: DataView,
	directory: PersistDirectory,
	persistId: number,
): PptRecord | undefined {
	const offset = directory.get(persistId);
	if (offset === undefined) {
		return undefined;
	}
	const rec = readRecordOrThrow(view, offset);
	if (rec.recLen > view.byteLength - rec.dataOffset) throw new PptParseError('Persist object payload out of bounds');
	return rec;
}
