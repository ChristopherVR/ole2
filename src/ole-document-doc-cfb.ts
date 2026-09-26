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
import { parseOle2 } from './ole2-parser-read.js';
import { ENTRY_TYPE_ROOT, ENTRY_TYPE_STREAM } from './ole2-parser-types.js';
import { buildOle2 } from './ole2-parser-write.js';

export interface DocCfbUnwrap {
	/** The `WordDocument` stream's bytes. */
	wordDocBytes: Uint8Array;
	/** The `0Table`/`1Table` stream `readDocFib` selected. */
	tableStreamName: '0Table' | '1Table';
	tableBytes: Uint8Array;
	/** Rebuild the full compound file with edited `WordDocument`/table stream bytes, preserving every other stream and the root CLSID. */
	rewrap: (editedWordDocBytes: Uint8Array, editedTableBytes: Uint8Array) => Uint8Array;
}

/** Unwrap a `.doc` payload's CFB container. Returns `undefined` if it is not a readable `WordDocument` CFB payload. */
export function unwrapDocBytes(bytes: Uint8Array): DocCfbUnwrap | undefined {
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
		const buffer = bytes.buffer.slice(
			bytes.byteOffset,
			bytes.byteOffset + bytes.byteLength,
		) as ArrayBuffer;
		const ole = parseOle2(buffer);
		const wordDocBytes = ole.getStream('WordDocument');
		if (!wordDocBytes) {
			return undefined;
		}
		const fib = readDocFib(wordDocBytes);
		const tableBytes = ole.getStream(fib.tableStreamName);
		if (!tableBytes) {
			return undefined;
		}
		const rootClsid = ole.entries.find((e) => e.type === ENTRY_TYPE_ROOT)?.clsid;
		return {
			wordDocBytes,
			tableStreamName: fib.tableStreamName,
			tableBytes,
			rewrap: (editedWordDocBytes, editedTableBytes) => {
				const streams = new Map<string, Uint8Array>();
				for (const entry of ole.entries) {
					if (entry.type !== ENTRY_TYPE_STREAM) {
						continue;
					}
					const data =
						entry.name === 'WordDocument'
							? editedWordDocBytes
							: entry.name === fib.tableStreamName
								? editedTableBytes
								: ole.getStream(entry.name);
					if (data) {
						streams.set(entry.name, data);
					}
				}
				return new Uint8Array(buildOle2(streams, rootClsid));
			},
		};
	} catch {
		return undefined;
	}
}
