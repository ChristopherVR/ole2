import { Ole2DocumentBase, UnsupportedOle2EditError, type Ole2DocumentCapabilities } from './ole2-document-base.js';
import type { Ole2DirectoryEntry } from './ole2-parser-types.js';
import { readCompoundFileStream } from './ole2-stream-edit.js';
import { resizeCompoundFileStream } from './ole2-stream-resize.js';
import { buildCompoundFile, listCompoundFile, type CompoundFileNode } from './ole2-tree-write.js';
import { CfbStorage, CfbStream, type CfbEntry, type CfbEntryAccess } from './cfb-storage.js';
import type { Ole2ModelDiagnostic } from './ole2-document.js';

const key = (path: readonly string[]) => JSON.stringify(path.map(name => name.toUpperCase()));
const within = (path: readonly string[], parent: readonly string[]) => path.length >= parent.length &&
	parent.every((name, index) => name.toUpperCase() === path[index]!.toUpperCase());
const fail = (reason: string): never => { throw new UnsupportedOle2EditError(reason); };
function validateName(name: string): void {
	if (typeof name !== 'string' || !name.length || name.length > 31 || /[\x00\\/:!]/.test(name) || name.toUpperCase().length !== name.length)
		fail('invalid-entry-name');
}
function validatePath(path: readonly string[]): void {
	if (!Array.isArray(path)) fail('invalid-entry-path');
	for (const name of path) validateName(name);
}
interface EntryReference { path: readonly string[]; live: boolean }

/** CFB edits do not repair Office records. Structural edits rebuild canonical
 * allocation using the original version, preserving reachable stream bytes
 * and supported metadata, rather than physical sectors or padding. */
export class CfbDocument extends Ole2DocumentBase {
	#references = new Map<string, EntryReference>();
	#editingStructure = false;
	get kind(): 'cfb' { return 'cfb'; }
	get capabilities(): Ole2DocumentCapabilities { return CFB_CAPABILITIES; }
	declare readonly detectedFormat: 'doc' | 'xls' | 'ppt' | 'vsd' | undefined;
	declare readonly diagnostics: readonly Ole2ModelDiagnostic[];
	constructor(input: Uint8Array, detectedFormat?: 'doc' | 'xls' | 'ppt' | 'vsd', diagnostics: readonly Ole2ModelDiagnostic[] = []) {
		super(input);
		Object.defineProperty(this, 'detectedFormat', { value: detectedFormat, enumerable: true, writable: false });
		Object.defineProperty(this, 'diagnostics', { value: Object.freeze(diagnostics.map(item => Object.freeze({ ...item }))), enumerable: true, writable: false });
	}
	get root(): CfbStorage { return this.storage([]); }
	storage(path: readonly string[]): CfbStorage {
		const reference = this.#reference(path, 'storage');
		const access = this.#access(reference);
		const childPath = (name: string) => { this.#entry(reference); validateName(name); return [...reference.path, name]; };
		return new CfbStorage({
			...access,
			children: () => {
				const parent = this.#entry(reference);
				return Object.freeze(this.entries.filter(entry => entry.parentId === parent.id).map(entry => this.#handle(entry.path!)));
			},
			storage: name => this.storage(childPath(name)),
			stream: name => this.stream(childPath(name)),
			createStorage: name => this.createStorage(childPath(name)),
			createStream: (name, bytes) => this.createStream(childPath(name), bytes),
		});
	}
	/** Existing stream access retains the preserving resize behavior. */
	stream(path: readonly string[]): CfbStream {
		const reference = this.#reference(path, 'stream');
		if (!readCompoundFileStream(this.getBytes(), reference.path)) fail('stream-not-found');
		const stream = new CfbStream(reference.path, () => {
			this.#entry(reference);
			return readCompoundFileStream(this.getBytes(), reference.path);
		}, bytes => {
			if (this.#editingStructure) fail('reentrant-cfb-edit');
			this.#entry(reference);
			const result = resizeCompoundFileStream(this.getBytes(), reference.path, bytes);
			if (!result.ok) fail(result.reason);
			this.commitBytes(result.bytes);
		}, this.#access(reference));
		Object.freeze(stream);
		return stream;
	}
	createStorage(path: readonly string[]): CfbStorage {
		return this.#structure(() => this.storage(this.#create(path)));
	}
	createStream(path: readonly string[], bytes: Uint8Array = new Uint8Array()): CfbStream {
		return this.#structure(() => {
			if (!(bytes instanceof Uint8Array)) fail('invalid-stream-bytes');
			return this.stream(this.#create(path, new Uint8Array(bytes)));
		});
	}
	/** Recursively deletes a storage, or deletes a single stream. */
	delete(path: readonly string[]): void {
		this.#structure(() => this.#delete(path));
	}
	#delete(path: readonly string[]): void {
		const source = this.#requireEntry(path);
		if (!source.path!.length) fail('root-edit-not-supported');
		const removed = source.path!;
		this.#rebuild(nodes => nodes.filter(node => !within(node.path, removed)));
		for (const [id, reference] of this.#references) if (within(reference.path, removed)) {
			reference.live = false; this.#references.delete(id);
		}
	}
	rename(path: readonly string[], name: string): CfbEntry {
		return this.#structure(() => {
			validateName(name);
			const source = this.#requireEntry(path);
			if (!source.path!.length) fail('root-edit-not-supported');
			return this.#move(source.path!, [...source.path!.slice(0, -1), name]);
		});
	}
	/** Moves an entry/subtree to a full destination path. Parent must exist. */
	move(path: readonly string[], destination: readonly string[]): CfbEntry {
		return this.#structure(() => this.#move(path, destination));
	}
	#move(path: readonly string[], destination: readonly string[]): CfbEntry {
		const source = this.#requireEntry(path), from = source.path!;
		validatePath(destination);
		if (!from.length || !destination.length) fail('root-edit-not-supported');
		const parent = this.#requireStorage(destination.slice(0, -1));
		const to = Object.freeze([...parent.path!, destination.at(-1)!]);
		if (from.length === to.length && from.every((name, index) => name === to[index])) return this.#handle(from);
		if (source.type === 1 && within(parent.path!, from)) fail('move-into-descendant');
		const occupied = this.findEntry(to);
		if (occupied && occupied.id !== source.id) fail('entry-already-exists');
		this.#rebuild(nodes => nodes.map(node => within(node.path, from) ? { ...node, path: [...to, ...node.path.slice(from.length)] } : node));
		const moved = [...this.#references.values()].filter(reference => within(reference.path, from));
		for (const reference of moved) this.#references.delete(key(reference.path));
		for (const reference of moved) {
			reference.path = Object.freeze([...to, ...reference.path.slice(from.length)]);
			this.#references.set(key(reference.path), reference);
		}
		return this.#handle(to);
	}
	#create(path: readonly string[], data?: Uint8Array): readonly string[] {
		validatePath(path);
		if (!path.length) fail('root-edit-not-supported');
		const parent = this.#requireStorage(path.slice(0, -1));
		const canonical = [...parent.path!, path.at(-1)!];
		if (this.findEntry(canonical)) fail('entry-already-exists');
		this.#rebuild(nodes => [...nodes, { path: canonical, data }]);
		return canonical;
	}
	#requireEntry(path: readonly string[]): Ole2DirectoryEntry {
		validatePath(path);
		return this.findEntry(path) ?? fail('entry-not-found');
	}
	#requireStorage(path: readonly string[]): Ole2DirectoryEntry {
		const entry = this.#requireEntry(path);
		if (entry.type !== 1 && entry.type !== 5) fail('not-a-storage');
		return entry;
	}
	#reference(path: readonly string[], kind: 'storage' | 'stream'): EntryReference {
		// Existing stream lookup retains the raw reader's accepted names; writer
		// name restrictions apply only when changing directory structure.
		if (!Array.isArray(path) || path.some(name => typeof name !== 'string')) fail('invalid-entry-path');
		const entry = this.findEntry(path) ?? fail(`${kind}-not-found`);
		if (kind === 'stream' ? entry.type !== 2 : entry.type !== 1 && entry.type !== 5) fail(`not-a-${kind}`);
		const id = key(entry.path!);
		let reference = this.#references.get(id);
		if (!reference) { reference = { path: Object.freeze([...entry.path!]), live: true }; this.#references.set(id, reference); }
		return reference;
	}
	#entry(reference: EntryReference): Ole2DirectoryEntry {
		if (!reference.live) fail('stale-entry');
		return this.findEntry(reference.path) ?? fail('stale-entry');
	}
	#handle(path: readonly string[]): CfbEntry {
		return this.#requireEntry(path).type === 2 ? this.stream(path) : this.storage(path);
	}
	#access(reference: EntryReference): CfbEntryAccess {
		return {
			path: () => reference.path,
			entry: () => this.#entry(reference),
			delete: () => { this.#entry(reference); this.delete(reference.path); },
			rename: name => { this.#entry(reference); this.rename(reference.path, name); },
			move: (destination, name) => {
				this.#structure(() => {
					this.#entry(reference); validatePath(destination);
					this.#move(reference.path, [...destination, name ?? reference.path.at(-1)!]);
				});
			},
		};
	}
	#rebuild(edit: (nodes: CompoundFileNode[]) => CompoundFileNode[]): void {
		const bytes = this.getBytes();
		// Refuse repair implicitly: unreachable entries/padding cannot be retained
		// by this writer, and every reachable stream must be readable.
		if (this.warnings.length) fail('structural-irregular-container');
		try {
			const { nodes, options } = listCompoundFile(bytes, { strict: true });
			const candidate = buildCompoundFile(edit(nodes), options);
			listCompoundFile(candidate, { strict: true });
			this.commitBytes(candidate);
		} catch (error) {
			throw new UnsupportedOle2EditError(`structural-rebuild-failed: ${error instanceof Error ? error.message : String(error)}`);
		}
	}
	#structure<T>(operation: () => T): T {
		if (this.#editingStructure) fail('reentrant-cfb-edit');
		this.#editingStructure = true;
		try { return operation(); }
		finally { this.#editingStructure = false; }
	}
}

const CFB_CAPABILITIES = Object.freeze({
	read: Object.freeze(['directory', 'stream-by-path', 'storage-hierarchy']),
	write: Object.freeze(['existing-stream-bytes', 'v3-create-delete-rename-move', 'v4-create-delete-rename-move']),
	limitations: Object.freeze([
		'Container editing does not update Office format records.',
		'Structural edits rebuild allocation using the original v3/v4 version; physical layout and padding are not preserved.',
		'Irregular containers are refused for structural edits.',
		'Preserving resize cannot expand external DIFAT or resize v4 mini streams.',
	]),
});
