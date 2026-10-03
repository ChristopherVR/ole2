/* Adapted from ChristopherVR/pptx-viewer packages/core/src/core/utils, Apache-2.0. Original source: https://github.com/ChristopherVR/pptx-viewer. */
import { readDocFib } from './ole-document-doc-fib.js';
/**
 * Unwrap/rewrap the OLE2 (MS-CFB) container around a legacy `.doc` (Word
 * 97-2003) embedding's `WordDocument` and table (`0Table`/`1Table`) streams.
 *
 * Mirrors `ole-sheet-xls-cfb.ts`'s `.xls` pattern exactly, generalised to two
 * streams that must be edited together (the piece table lives in the table
 * stream, but its FKP pages and text bytes live in `WordDocument`; a
 * paragraph edit touches both). Every other stream (`\x01CompObj`,
 * `\x05SummaryInformation`, etc.) and the root entry's CLSID round-trip
 * byte-for-byte.
 *
 * @module ole-document-doc-cfb
 */
import { readCompoundFileStream, replaceCompoundFileStream } from './ole2-stream-edit.js';
import { resizeCompoundFileStream } from './ole2-stream-resize.js';

/** A container edit was refused rather than rebuilding and losing opaque data. */
export class DocCfbRewriteError extends Error {
	constructor(message: string) {
		super(message);
		this.name = 'DocCfbRewriteError';
	}
}

export interface DocCfbUnwrap {
	/** The `WordDocument` stream's bytes. */
	wordDocBytes: Uint8Array;
	/** The `0Table`/`1Table` stream `readDocFib` selected. */
	tableStreamName: '0Table' | '1Table';
	tableBytes: Uint8Array;
	/** Both DOC streams can grow through the preserving regular-v3 stream resizer. */
	canRewrite?: boolean;
	/** Patch or resize the DOC streams while retaining the original hierarchy, metadata and other streams. Throws if preservation is unsupported. */
	rewrap: (editedWordDocBytes: Uint8Array, editedTableBytes: Uint8Array) => Uint8Array;
}

/** Container capabilities known after parsing; callers can still construct the
 * original public DocCfbUnwrap interface without the additive capability field. */
export interface ParsedDocCfbUnwrap extends DocCfbUnwrap {
	canRewrite: boolean;
}

/** Unwrap a `.doc` payload's CFB container. Returns `undefined` if it is not a readable `WordDocument` CFB payload. */
export function unwrapDocBytes(bytes: Uint8Array): ParsedDocCfbUnwrap | undefined {
	if (
		bytes.length < 8 ||
		bytes[0] !== 0xd0 ||
		bytes[1] !== 0xcf ||
		bytes[2] !== 0x11 ||
		bytes[3] !== 0xe0 ||
		bytes[4] !== 0xa1 ||
		bytes[5] !== 0xb1 ||
		bytes[6] !== 0x1a ||
		bytes[7] !== 0xe1
	) {
		return undefined;
	}
	try {
		const wordDocBytes = readCompoundFileStream(bytes, ['WordDocument']);
		if (!wordDocBytes) {
			return undefined;
		}
		const fib = readDocFib(wordDocBytes);
		const tableBytes = readCompoundFileStream(bytes, [fib.tableStreamName]);
		if (!tableBytes) {
			return undefined;
		}
		const header = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
		const cutoff = header.getUint32(0x38, true);
		const canRewrite = header.getUint16(0x1a, true) === 3 && header.getUint32(0x48, true) === 0 &&
			header.getUint32(0x2c, true) <= 109 && (cutoff === 0 || cutoff === 4096) &&
			wordDocBytes.length >= cutoff && tableBytes.length >= cutoff;
		return {
			wordDocBytes,
			tableStreamName: fib.tableStreamName,
			tableBytes,
			canRewrite,
			rewrap: (editedWordDocBytes, editedTableBytes) => {
				if (editedWordDocBytes.length === wordDocBytes.length && editedTableBytes.length === tableBytes.length) {
					const patchedWord = replaceCompoundFileStream(bytes, ['WordDocument'], editedWordDocBytes);
					if (patchedWord === bytes) throw new DocCfbRewriteError('Cannot safely patch WordDocument');
					if (editedTableBytes.every((value, i) => value === tableBytes[i])) return patchedWord;
					const patchedTable = replaceCompoundFileStream(patchedWord, [fib.tableStreamName], editedTableBytes);
					if (patchedTable === patchedWord) throw new DocCfbRewriteError('Cannot safely patch DOC table stream');
					return patchedTable;
				}
				const word = resizeCompoundFileStream(bytes, ['WordDocument'], editedWordDocBytes);
				if (!word.ok) throw new DocCfbRewriteError(`Cannot resize WordDocument: ${word.reason}`);
				const table = resizeCompoundFileStream(word.bytes, [fib.tableStreamName], editedTableBytes);
				if (!table.ok) throw new DocCfbRewriteError(`Cannot resize DOC table stream: ${table.reason}`);
				return table.bytes;
			},
		};
	} catch {
		return undefined;
	}
}
