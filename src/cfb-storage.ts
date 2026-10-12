import type { Ole2DirectoryEntry } from './ole2-parser-types.js';
import { UnsupportedOle2EditError } from './ole2-document-base.js';

/** Document-owned operations used by storage and stream handles. */
export interface CfbEntryAccess {
	path(): readonly string[];
	entry(): Ole2DirectoryEntry;
	delete(): void;
	rename(name: string): void;
	move(destination: readonly string[], name?: string): void;
}

export interface CfbStorageAccess extends CfbEntryAccess {
	children(): readonly CfbEntry[];
	storage(name: string): CfbStorage;
	stream(name: string): CfbStream;
	createStorage(name: string): CfbStorage;
	createStream(name: string, bytes: Uint8Array): CfbStream;
}

export type CfbEntry = CfbStorage | CfbStream;

/** A live storage handle. Paths follow moves; deleted handles remain stale. */
export class CfbStorage {
	readonly kind = 'storage';
	#access: CfbStorageAccess;
	constructor(access: CfbStorageAccess) { this.#access = access; Object.freeze(this); }
	get path(): readonly string[] { return this.#access.path(); }
	get name(): string { return this.entry.name; }
	get entry(): Ole2DirectoryEntry { return this.#access.entry(); }
	get children(): readonly CfbEntry[] { return this.#access.children(); }
	storage(name: string): CfbStorage { return this.#access.storage(name); }
	stream(name: string): CfbStream { return this.#access.stream(name); }
	createStorage(name: string): CfbStorage { return this.#access.createStorage(name); }
	createStream(name: string, bytes: Uint8Array = new Uint8Array()): CfbStream { return this.#access.createStream(name, bytes); }
	/** Deletes this storage and all descendants. The root cannot be deleted. */
	delete(): void { this.#access.delete(); }
	rename(name: string): void { this.#access.rename(name); }
	/** Destination is an absolute storage path; [] selects the root. */
	move(destination: readonly string[], name?: string): void { this.#access.move(destination, name); }
}

export class CfbStream {
	readonly kind = 'stream';
	#path: readonly string[];
	#access?: CfbEntryAccess;
	/** The original three-argument constructor remains supported. */
	constructor(path: readonly string[], private readonly read: () => Uint8Array | undefined,
		private readonly write: (bytes: Uint8Array) => void, access?: CfbEntryAccess) {
		this.#path = Object.freeze([...path]); this.#access = access;
	}
	get path(): readonly string[] { return this.#access?.path() ?? this.#path; }
	get name(): string { return this.path.at(-1)!; }
	get entry(): Ole2DirectoryEntry { return this.#requiredAccess().entry(); }
	get bytes(): Uint8Array { const value = this.read(); if (!value) throw new UnsupportedOle2EditError('stale-stream'); return new Uint8Array(value); }
	set bytes(value: Uint8Array) { if (!(value instanceof Uint8Array)) throw new UnsupportedOle2EditError('invalid-stream-bytes'); this.write(new Uint8Array(value)); }
	delete(): void { this.#requiredAccess().delete(); }
	rename(name: string): void { this.#requiredAccess().rename(name); }
	/** Destination is an absolute storage path; [] selects the root. */
	move(destination: readonly string[], name?: string): void { this.#requiredAccess().move(destination, name); }
	#requiredAccess(): CfbEntryAccess {
		if (!this.#access) throw new UnsupportedOle2EditError('unowned-stream');
		return this.#access;
	}
}
