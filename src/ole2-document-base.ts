/** Shared ownership and transactional serialization for editable legacy models. */
import { parseOle2 as parseContainer } from './ole2-parser-read.js';
import type { Ole2DirectoryEntry, Ole2File, Ole2ParseWarning } from './ole2-parser-types.js';

const copyOf = (bytes: Uint8Array | undefined) => bytes === undefined ? undefined : new Uint8Array(bytes);
const detachEntry = (entry: Ole2DirectoryEntry): Ole2DirectoryEntry => ({
 ...entry,
 clsid: new Uint8Array(entry.clsid),
 path: entry.path && [...entry.path],
 created: entry.created && new Date(entry.created),
 modified: entry.modified && new Date(entry.modified),
});

export type Ole2DocumentKind = 'cfb' | 'doc' | 'xls' | 'ppt' | 'vsd';
export interface Ole2DocumentCapabilities {
 readonly read: readonly string[];
 readonly write: readonly string[];
 readonly limitations: readonly string[];
}
export class Ole2DocumentError extends Error {
 constructor(readonly code: 'format-mismatch' | 'ambiguous-format' | 'invalid-format' | 'unsupported-model', message: string) {
  super(message); this.name = 'Ole2DocumentError';
 }
}
export class UnsupportedOle2EditError extends Error {
 constructor(readonly reason: string) { super(`Unsupported OLE2 edit: ${reason}`); this.name = 'UnsupportedOle2EditError'; }
}

/** Models own their bytes. Inspection arrays and returned streams are detached
 * snapshots; supported model setters are the only document mutation path. */
export abstract class Ole2DocumentBase implements Ole2File {
 abstract readonly kind: Ole2DocumentKind;
 abstract readonly capabilities: Ole2DocumentCapabilities;
 #bytes: Uint8Array;
 #container: Ole2File;
 #dirty = false;
 #revision = 0;
 constructor(input: Uint8Array) {
  this.#bytes = new Uint8Array(input);
  this.#container = parseContainer(this.#bytes.buffer as ArrayBuffer);
 }
 get dirty(): boolean { return this.#dirty; }
 get revision(): number { return this.#revision; }
 get entries(): Ole2DirectoryEntry[] {
  return this.#container.entries.map(detachEntry);
 }
 get warnings(): Ole2ParseWarning[] { return this.#container.warnings.map(warning => ({...warning})); }
 readonly #getStream = (name: string): Uint8Array | undefined => copyOf(this.#container.getStream(name));
 readonly #getStreamByPath = (path: string | readonly string[]): Uint8Array | undefined => copyOf(this.#container.getStreamByPath(path));
 readonly #getStreamById = (id: number): Uint8Array | undefined => copyOf(this.#container.getStreamById(id));
 readonly #findEntry = (path: string | readonly string[]): Ole2DirectoryEntry | undefined => {
  const entry = this.#container.findEntry(path);
  return entry && detachEntry(entry);
 };
 /** Retain the original detachable callback API while observing current state. */
 get getStream(): Ole2File['getStream'] { return this.#getStream; }
 get getStreamByPath(): Ole2File['getStreamByPath'] { return this.#getStreamByPath; }
 get getStreamById(): Ole2File['getStreamById'] { return this.#getStreamById; }
 get findEntry(): Ole2File['findEntry'] { return this.#findEntry; }
 /** Isolated working input: codecs cannot accidentally mutate current state. */
 protected getBytes(): Uint8Array { return new Uint8Array(this.#bytes); }
 /** Validate a candidate before changing any owned state. No-op stays clean. */
 protected commitBytes(candidate: Uint8Array): void {
  const next = new Uint8Array(candidate);
  const container = parseContainer(next.buffer as ArrayBuffer);
  if(next.length === this.#bytes.length && next.every((value,index) => value === this.#bytes[index])) return;
  this.#bytes = next; this.#container = container; this.#dirty = true; this.#revision++;
 }
 /** Serialization retains opaque bytes and returns an independent copy. It
 * does not reset dirty state or imply unsupported content can be reconstructed. */
 serialize(): Uint8Array { return new Uint8Array(this.#bytes); }
}
