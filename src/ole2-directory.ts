/**
 * MS-CFB directory decoding: slot records, red-black tree resolution into
 * storage paths, and name/path lookup.
 *
 * Reference: [MS-CFB] 2.6 Compound File Directory Sectors
 *
 * @module ole2-directory
 */

import type { Ole2DirectoryEntry, Ole2ParseWarning } from './ole2-parser-types.js';
import { fileTimeDate } from './ole2-filetime.js';
import {
	DIR_ENTRY_SIZE,
	ENTRY_TYPE_EMPTY,
	ENTRY_TYPE_ROOT,
	ENTRY_TYPE_STORAGE,
	NOSTREAM,
	Ole2ParseError,
} from './ole2-parser-types.js';

interface Slot {
	id: number;
	type: number;
	nameValid: boolean;
	entry: Ole2DirectoryEntry;
}

function decodeSlot(dir: Uint8Array, id: number, majorVersion: number, warnings: Ole2ParseWarning[]): Slot {
	const view = new DataView(dir.buffer, dir.byteOffset + id * DIR_ENTRY_SIZE, DIR_ENTRY_SIZE);
	const nameLen = view.getUint16(64, true);
	const type = view.getUint8(66);
	const nameValid = nameLen >= 2 && nameLen <= 64 && nameLen % 2 === 0 && view.getUint16(nameLen - 2, true) === 0;
	let name = '';
	if (nameValid) for (let j = 0; j < nameLen - 2; j += 2) name += String.fromCharCode(view.getUint16(j, true));
	const link = (offset: number) => {
		const value = view.getUint32(offset, true);
		return value === NOSTREAM ? -1 : value;
	};
	const sizeLow = view.getUint32(120, true);
	const sizeHigh = view.getUint32(124, true);
	let size = sizeLow;
	if (majorVersion === 4) size = sizeLow + sizeHigh * 0x100000000;
	else if (sizeHigh !== 0 && type !== ENTRY_TYPE_EMPTY) {
		// [MS-CFB] 2.6.3: old v3 writers left these bits uninitialized; readers should ignore them.
		warnings.push({ code: 'size-high-bits', message: `Ignored v3 stream-size high bits on directory entry ${id}` });
	}
	return {
		id,
		type,
		nameValid,
		entry: {
			id,
			name,
			type,
			color: view.getUint8(67),
			leftSiblingId: link(68),
			rightSiblingId: link(72),
			childId: link(76),
			clsid: new Uint8Array(view.buffer.slice(view.byteOffset + 80, view.byteOffset + 96)),
			stateBits: view.getUint32(96, true),
			created: fileTimeDate(view.getBigUint64(100, true)),
			modified: fileTimeDate(view.getBigUint64(108, true)),
			createdFileTime: view.getBigUint64(100, true),
			modifiedFileTime: view.getBigUint64(108, true),
			startSector: view.getUint32(116, true),
			size,
			parentId: -1,
		},
	};
}

/**
 * Decode every directory slot and resolve storage paths. Stale slots that no
 * storage reaches are kept (lenient) or rejected (strict); a reachable slot
 * with an undecodable name always fails.
 */
export function readDirectory(
	dir: Uint8Array,
	majorVersion: number,
	strict: boolean,
	warnings: Ole2ParseWarning[],
): Ole2DirectoryEntry[] {
	const count = Math.floor(dir.length / DIR_ENTRY_SIZE);
	const slots: Slot[] = [];
	for (let id = 0; id < count; id++) slots.push(decodeSlot(dir, id, majorVersion, warnings));
	const roots = slots.filter((slot) => slot.type === ENTRY_TYPE_ROOT);
	if (roots.length !== 1 || !roots[0]!.nameValid) throw new Ole2ParseError('Invalid root directory entry');
	const root = roots[0]!;
	root.entry.path = [];

	const reached = new Set<number>([root.id]);
	const fail = (message: string) => {
		if (strict) throw new Ole2ParseError(message);
		warnings.push({ code: 'directory-link', message });
	};
	// Iterative walk: hostile trees can be deep enough to overflow the JS stack.
	const pending: Array<{ id: number; parent: Slot }> = [];
	const push = (id: number, parent: Slot, from: number) => {
		if (id === -1) return;
		const target = slots[id];
		if (!target || target.type === ENTRY_TYPE_EMPTY || target.type === ENTRY_TYPE_ROOT) {
			fail(`Directory entry ${from} links to invalid entry ${id}`);
			return;
		}
		if (reached.has(id)) {
			fail(`Cyclic directory tree at entry ${id}`);
			return;
		}
		reached.add(id);
		pending.push({ id, parent });
	};
	push(root.entry.childId, root, root.id);
	while (pending.length) {
		const { id, parent } = pending.pop()!;
		const slot = slots[id]!;
		if (!slot.nameValid) throw new Ole2ParseError('Invalid directory entry name');
		slot.entry.parentId = parent.id;
		slot.entry.path = [...parent.entry.path!, slot.entry.name];
		push(slot.entry.leftSiblingId, parent, id);
		push(slot.entry.rightSiblingId, parent, id);
		if (slot.type === ENTRY_TYPE_STORAGE) push(slot.entry.childId, slot, id);
	}

	const entries: Ole2DirectoryEntry[] = [];
	for (const slot of slots) {
		if (slot.type === ENTRY_TYPE_EMPTY) continue;
		if (!reached.has(slot.id)) {
			if (!slot.nameValid) {
				if (strict) throw new Ole2ParseError('Invalid directory entry name');
				warnings.push({ code: 'invalid-unreachable-entry', message: `Skipped undecodable stale directory entry ${slot.id}` });
				continue;
			}
			if (strict) throw new Ole2ParseError(`Directory entry ${slot.id} is not reachable from the root`);
			warnings.push({ code: 'unreachable-entry', message: `Directory entry ${slot.id} (${slot.entry.name}) is not reachable from the root` });
		}
		entries.push(slot.entry);
	}
	return entries;
}

/** [MS-CFB] 2.6.4 name equality: uppercase UTF-16 code-unit comparison. */
function sameName(a: string, b: string): boolean {
	return a.length === b.length && a.toUpperCase() === b.toUpperCase();
}

/** Strip the leading property-set/OLE control prefix (\u0001-\u0006). */
function bareName(name: string): string {
	return name.replace(/^[\u0001-\u0006]/, '');
}

/** Split a `/`-separated path; an optional leading `/` or root name is ignored. */
export function normalizePath(path: string | readonly string[], rootName: string): string[] {
	const parts = typeof path === 'string' ? path.split('/').filter((part) => part !== '') : [...path];
	if (parts.length && sameName(parts[0]!, rootName) && typeof path === 'string') parts.shift();
	return parts;
}

/** Pick one child of `parentId` by name; ambiguous matches resolve to nothing. */
function child(entries: readonly Ole2DirectoryEntry[], parentId: number, name: string): Ole2DirectoryEntry | undefined {
	const children = entries.filter((entry) => entry.parentId === parentId && entry.path !== undefined);
	const pick = (matches: Ole2DirectoryEntry[]) => (matches.length === 1 ? matches[0] : undefined);
	const exact = children.filter((entry) => entry.name === name);
	if (exact.length) return pick(exact);
	const folded = children.filter((entry) => sameName(entry.name, name));
	if (folded.length) return pick(folded);
	return pick(children.filter((entry) => sameName(bareName(entry.name), bareName(name))));
}

/** Resolve a storage path to its directory entry. */
export function findDirectoryEntry(
	entries: readonly Ole2DirectoryEntry[],
	path: string | readonly string[],
): Ole2DirectoryEntry | undefined {
	const root = entries.find((entry) => entry.type === ENTRY_TYPE_ROOT);
	if (!root) return undefined;
	let current: Ole2DirectoryEntry | undefined = root;
	for (const name of normalizePath(path, root.name)) {
		if (current.type !== ENTRY_TYPE_ROOT && current.type !== ENTRY_TYPE_STORAGE) return undefined;
		current = child(entries, current.id, name);
		if (!current) return undefined;
	}
	return current;
}

/**
 * Legacy single-name lookup: root-level children first (exact, then
 * case-insensitive), then the first exact match anywhere in slot order,
 * which keeps names that only exist in nested storages addressable.
 */
export function findStreamByName(
	entries: readonly Ole2DirectoryEntry[],
	name: string,
	isReadable: (entry: Ole2DirectoryEntry) => boolean,
): Ole2DirectoryEntry | undefined {
	const root = entries.find((entry) => entry.type === ENTRY_TYPE_ROOT);
	if (root && root.name === name) return root;
	const top = entries.filter((entry) => entry.parentId === root?.id && isReadable(entry));
	return (
		top.find((entry) => entry.name === name) ??
		top.find((entry) => sameName(entry.name, name)) ??
		entries.find((entry) => entry.name === name && isReadable(entry))
	);
}
