/** Shared ownership and transactional serialization for editable legacy models. */
import { parseOle2 as parseContainer } from './ole2-parser-read.js';
import type { Ole2DirectoryEntry, Ole2File } from './ole2-parser-types.js';

export type Ole2DocumentKind = 'cfb' | 'doc' | 'xls' | 'ppt';
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
  return this.#container.entries.map(entry => ({...entry, clsid: new Uint8Array(entry.clsid)}));
 }
 readonly #getStream = (name: string): Uint8Array | undefined => {
  const bytes = this.#container.getStream(name);
  return bytes === undefined ? undefined : new Uint8Array(bytes);
 };
 /** Retain the original detachable callback API while observing current state. */
 get getStream(): Ole2File['getStream'] { return this.#getStream; }
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
