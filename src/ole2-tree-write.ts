/**
 * Hierarchical MS-CFB writer, listing and repair.
 *
 * `buildOle2` stays the flat writer whose exact layout PowerPoint was
 * validated against. This module adds what a general compound-file library
 * needs on top: nested storages, per-entry CLSIDs, state bits and
 * timestamps, a `listCompoundFile` -> edit -> `buildCompoundFile` round trip
 * (add, delete or move entries by editing the node array), and
 * `repairCompoundFile`, which rewrites a leniently readable container into a
 * canonical one.
 *
 * Reference: [MS-CFB] Compound Binary File Format
 *
 * @module ole2-tree-write
 */

import {
	DIFSECT,
	DIR_ENTRY_SIZE,
	ENDOFCHAIN,
	ENTRY_TYPE_ROOT,
	ENTRY_TYPE_STORAGE,
	ENTRY_TYPE_STREAM,
	FATSECT,
	NOSTREAM,
	type Ole2ParseWarning,
	type Ole2ReadOptions,
} from './ole2-parser-types.js';
import { parseOle2 } from './ole2-parser-read.js';
import { compareDirEntryNames, encodeName, writeFatChain, writeFatRun, writeInt32Sectors } from './ole2-parser-write-helpers.js';
import { sizeFatSectors, writeDifatSectors, writeHeader, writeStreamSectors } from './ole2-parser-write-serialize.js';

/** One storage or stream below the root. A node with `data` is a stream. */
export interface CompoundFileNode {
	/** Names from the root down, e.g. `['ObjectPool', '_1234', 'Workbook']`. */
	path: readonly string[];
	/** Stream bytes. Omit for a storage. */
	data?: Uint8Array;
	clsid?: Uint8Array;
	stateBits?: number;
	created?: Date;
	modified?: Date;
}

export interface CompoundFileBuildOptions {
	rootClsid?: Uint8Array;
	rootStateBits?: number;
	rootModified?: Date;
	/** Streams smaller than this go to the mini stream. Default 4096; 0 disables it. */
	miniStreamCutoff?: number;
}

const SECTOR = 512;
const MINI_SECTOR = 64;
const FILETIME_UNIX_EPOCH = 116444736000000000n;

interface Node {
	name: string;
	type: number;
	data?: Uint8Array;
	clsid?: Uint8Array;
	stateBits: number;
	created?: Date;
	modified?: Date;
	children: Node[];
	id: number;
	left: number;
	right: number;
	child: number;
	color: number;
	start: number;
}

function validateName(name: string): void {
	if (name.toUpperCase().length !== name.length) {
		throw new Error('CFB names with expanding uppercase mappings are unsupported');
	}
	if (!name.length || name.length > 31 || /[\x00\\/:!]/.test(name)) {
		throw new Error(`CFB name "${name}" must be 1-31 UTF-16 code units without NUL, \\, /, :, or !`);
	}
}

function writeFileTime(view: DataView, offset: number, date: Date | undefined): void {
	if (!date || Number.isNaN(date.getTime())) return;
	const ticks = BigInt(Math.round(date.getTime())) * 10000n + FILETIME_UNIX_EPOCH;
	if (ticks > 0n) view.setBigUint64(offset, ticks, true);
}

/** Balanced red-black sibling tree over name-sorted nodes (same shape as `buildDirectoryTree`). */
function linkSiblings(nodes: Node[]): number {
	if (!nodes.length) return NOSTREAM;
	let redLevel = 0;
	for (let remaining = nodes.length - 1; remaining >= 0; remaining = Math.floor(remaining / 2) - 1) redLevel++;
	const build = (low: number, high: number, level: number): number => {
		if (low > high) return NOSTREAM;
		const middle = Math.floor((low + high) / 2);
		const node = nodes[middle]!;
		node.left = build(low, middle - 1, level + 1);
		node.right = build(middle + 1, high, level + 1);
		node.color = level === redLevel ? 0 : 1;
		return node.id;
	};
	return build(0, nodes.length - 1, 0);
}

function buildTree(input: readonly CompoundFileNode[], options: CompoundFileBuildOptions): Node[] {
	if (options.rootClsid && options.rootClsid.length !== 16) throw new Error('Root CLSID must contain 16 bytes');
	const root: Node = {
		name: 'Root Entry', type: ENTRY_TYPE_ROOT, clsid: options.rootClsid, stateBits: options.rootStateBits ?? 0,
		modified: options.rootModified, children: [], id: 0, left: NOSTREAM, right: NOSTREAM, child: NOSTREAM, color: 1, start: ENDOFCHAIN,
	};
	const byKey = new Map<string, Node>([['', root]]);
	const keyOf = (path: readonly string[]) => path.map((name) => name.toUpperCase()).join('/');
	const ensure = (path: readonly string[]): Node => {
		const key = keyOf(path);
		const found = byKey.get(key);
		if (found) return found;
		const parent = ensure(path.slice(0, -1));
		if (parent.type === ENTRY_TYPE_STREAM) throw new Error(`"${path.slice(0, -1).join('/')}" is a stream, not a storage`);
		const name = path.at(-1)!;
		validateName(name);
		const node: Node = {
			name, type: ENTRY_TYPE_STORAGE, stateBits: 0, children: [], id: -1, left: NOSTREAM, right: NOSTREAM, child: NOSTREAM, color: 1, start: 0,
		};
		parent.children.push(node);
		byKey.set(key, node);
		return node;
	};
	const declared = new Set<string>();
	for (const item of input) {
		if (!item.path.length) throw new Error('CFB node paths must name at least one entry');
		const key = keyOf(item.path);
		if (declared.has(key)) throw new Error(`Duplicate CFB path "${item.path.join('/')}"`);
		declared.add(key);
		if (item.clsid && item.clsid.length !== 16) throw new Error('CLSID must contain 16 bytes');
		const node = ensure(item.path);
		if (item.data) {
			if (node.children.length) throw new Error(`"${item.path.join('/')}" is a storage, not a stream`);
			node.type = ENTRY_TYPE_STREAM;
			node.data = item.data;
		}
		node.clsid = item.clsid;
		node.stateBits = item.stateBits ?? 0;
		node.created = item.created;
		node.modified = item.modified;
	}
	// Directory IDs in depth-first order; each storage's children sorted per [MS-CFB] 2.6.4.
	const ordered: Node[] = [];
	const visit = (node: Node) => {
		node.id = ordered.length;
		ordered.push(node);
		node.children.sort((a, b) => compareDirEntryNames(a.name, b.name));
		for (const child of node.children) visit(child);
	};
	visit(root);
	for (const node of ordered) node.child = linkSiblings(node.children);
	return ordered;
}

/**
 * Build a v3 compound file from storages and streams. Parent storages are
 * created implicitly; list a storage explicitly to give it a CLSID, state
 * bits or timestamps, or to keep it when empty.
 */
export function buildCompoundFile(nodes: readonly CompoundFileNode[], options: CompoundFileBuildOptions = {}): Uint8Array {
	const cutoff = options.miniStreamCutoff ?? 0x1000;
	const ordered = buildTree(nodes, options);
	const streams = ordered.filter((node) => node.type === ENTRY_TYPE_STREAM);
	const mini = streams.filter((node) => node.data!.length < cutoff);
	const regular = streams.filter((node) => node.data!.length >= cutoff);

	let next = 0;
	const chains = new Map<Node, number[]>();
	const allocate = (count: number) => Array.from({ length: count }, () => next++);
	for (const node of regular) chains.set(node, allocate(Math.ceil(node.data!.length / SECTOR)));

	let miniSectors = 0;
	const miniChains = new Map<Node, number[]>();
	for (const node of mini) {
		const count = Math.ceil(node.data!.length / MINI_SECTOR);
		miniChains.set(node, Array.from({ length: count }, () => miniSectors++));
	}
	const container = new Uint8Array(miniSectors * MINI_SECTOR);
	for (const node of mini) miniChains.get(node)!.forEach((id, i) => {
		container.set(node.data!.subarray(i * MINI_SECTOR, (i + 1) * MINI_SECTOR), id * MINI_SECTOR);
	});
	const rootSectors = allocate(Math.ceil(container.length / SECTOR));
	const dirSectors = Math.ceil((ordered.length * DIR_ENTRY_SIZE) / SECTOR);
	const firstDir = next;
	next += dirSectors;
	const miniFatSectors = miniSectors ? Math.ceil((miniSectors * 4) / SECTOR) : 0;
	const firstMiniFat = miniFatSectors ? next : ENDOFCHAIN;
	next += miniFatSectors;
	const { numFATSectors, numDIFATSectors } = sizeFatSectors(next, SECTOR);
	const firstFat = next;
	next += numFATSectors;
	const firstDifat = next;
	next += numDIFATSectors;

	const fat = new Int32Array(numFATSectors * (SECTOR / 4)).fill(-1);
	for (const sectors of chains.values()) writeFatChain(fat, { start: sectors[0] ?? ENDOFCHAIN, sectors });
	writeFatChain(fat, { start: rootSectors[0] ?? ENDOFCHAIN, sectors: rootSectors });
	writeFatRun(fat, firstDir, dirSectors);
	if (miniFatSectors) writeFatRun(fat, firstMiniFat, miniFatSectors);
	for (let i = 0; i < numFATSectors; i++) fat[firstFat + i] = FATSECT;
	for (let i = 0; i < numDIFATSectors; i++) fat[firstDifat + i] = DIFSECT;

	const out = new Uint8Array((next + 1) * SECTOR);
	const view = new DataView(out.buffer);
	writeHeader(view, out, {
		numFATSectors, firstDirSector: firstDir, miniStreamCutoff: cutoff,
		firstMiniFATSector: firstMiniFat, numMiniFATSectors: miniFatSectors,
		firstFATSector: firstFat, firstDIFATSector: firstDifat, numDIFATSectors,
	});
	for (const [node, sectors] of chains) writeStreamSectors(out, node.data!, { start: sectors[0]!, sectors }, SECTOR);
	if (rootSectors.length) writeStreamSectors(out, container, { start: rootSectors[0]!, sectors: rootSectors }, SECTOR);
	if (miniFatSectors) {
		const miniFat = new Int32Array(miniFatSectors * (SECTOR / 4)).fill(-1);
		for (const sectors of miniChains.values()) writeFatChain(miniFat, { start: sectors[0] ?? ENDOFCHAIN, sectors });
		writeInt32Sectors(out, miniFat, firstMiniFat, miniFatSectors, SECTOR);
	}

	// Unused directory slots stay zero (empty, name length 0) except their links.
	const dirBase = (firstDir + 1) * SECTOR;
	for (let slot = ordered.length; slot < (dirSectors * SECTOR) / DIR_ENTRY_SIZE; slot++) {
		for (const field of [68, 72, 76]) view.setUint32(dirBase + slot * DIR_ENTRY_SIZE + field, NOSTREAM, true);
	}
	for (const node of ordered) {
		const at = dirBase + node.id * DIR_ENTRY_SIZE;
		out.set(encodeName(node.name), at);
		view.setUint16(at + 64, (node.name.length + 1) * 2, true);
		out[at + 66] = node.type;
		out[at + 67] = node.id === 0 ? 1 : node.color;
		view.setUint32(at + 68, node.id === 0 ? NOSTREAM : node.left, true);
		view.setUint32(at + 72, node.id === 0 ? NOSTREAM : node.right, true);
		view.setUint32(at + 76, node.type === ENTRY_TYPE_STREAM ? NOSTREAM : node.child, true);
		if (node.clsid) out.set(node.clsid, at + 80);
		view.setUint32(at + 96, node.stateBits >>> 0, true);
		// [MS-CFB] 2.6.3: the root entry's creation time must be zero.
		if (node.id !== 0) writeFileTime(view, at + 100, node.created);
		writeFileTime(view, at + 108, node.modified);
		let start = 0, size = 0;
		if (node.type === ENTRY_TYPE_ROOT) { start = rootSectors[0] ?? ENDOFCHAIN; size = container.length; }
		else if (node.type === ENTRY_TYPE_STREAM) {
			size = node.data!.length;
			start = (chains.get(node) ?? miniChains.get(node))![0] ?? ENDOFCHAIN;
		}
		view.setUint32(at + 116, start, true);
		view.setUint32(at + 120, size, true);
	}
	writeInt32Sectors(out, fat, firstFat, numFATSectors, SECTOR);
	writeDifatSectors(view, { firstFATSector: firstFat, numFATSectors, firstDIFATSector: firstDifat, numDIFATSectors, sectorSize: SECTOR });
	return out;
}

type ParsedFile = ReturnType<typeof parseOle2>;

/** Reachable storages and streams; unreadable streams throw or land in `dropped`. */
function collect(file: ParsedFile, dropUnreadable: boolean) {
	const root = file.entries.find((entry) => entry.type === ENTRY_TYPE_ROOT)!;
	const nodes: CompoundFileNode[] = [];
	const dropped: string[][] = [];
	for (const entry of file.entries) {
		if (!entry.path?.length || (entry.type !== ENTRY_TYPE_STORAGE && entry.type !== ENTRY_TYPE_STREAM)) continue;
		let data: Uint8Array | undefined;
		if (entry.type === ENTRY_TYPE_STREAM) {
			try {
				data = file.getStreamById(entry.id);
			} catch (error) {
				if (!dropUnreadable) throw error;
				dropped.push([...entry.path]);
				continue;
			}
		}
		nodes.push({
			path: [...entry.path], data, stateBits: entry.stateBits, created: entry.created, modified: entry.modified,
			clsid: entry.clsid.some((byte) => byte !== 0) ? entry.clsid : undefined,
		});
	}
	const options: CompoundFileBuildOptions = { rootClsid: root.clsid, rootStateBits: root.stateBits, rootModified: root.modified };
	return { nodes, options, dropped };
}

/**
 * Every reachable storage and stream with its bytes. Edit the array (push to
 * add, filter to delete, change `path` to move or rename) and pass it with
 * `options` to `buildCompoundFile`.
 */
export function listCompoundFile(input: Uint8Array | ArrayBuffer, readOptions: Ole2ReadOptions = {}): {
	nodes: CompoundFileNode[];
	options: CompoundFileBuildOptions;
	warnings: Ole2ParseWarning[];
} {
	const file = parseOle2(input, readOptions);
	const { nodes, options } = collect(file, false);
	return { nodes, options, warnings: file.warnings };
}

export interface CompoundFileRepairResult {
	bytes: Uint8Array;
	/** Irregularities the lenient reader accepted in the input. */
	warnings: Ole2ParseWarning[];
	/** Streams whose allocation could not be read without inventing bytes. */
	dropped: string[][];
}

/**
 * Rewrite a leniently readable container as a canonical v3 file: one root,
 * contiguous chains, balanced directory trees and exact header counts, with
 * names, CLSIDs, state bits and timestamps kept. Stale unreachable slots are
 * discarded. A stream whose bytes cannot be read throws unless
 * `dropUnreadable` is set, in which case it is omitted and listed in `dropped`.
 */
export function repairCompoundFile(input: Uint8Array | ArrayBuffer, options: { dropUnreadable?: boolean } = {}): CompoundFileRepairResult {
	const file = parseOle2(input);
	const { nodes, options: build, dropped } = collect(file, options.dropUnreadable === true);
	return { bytes: buildCompoundFile(nodes, build), warnings: file.warnings, dropped };
}
