/* Adapted from ChristopherVR/pptx-viewer under Apache-2.0. Original: packages/core/src/core/utils/ole-sheet-xls-cfb.ts. */
/**
 * Unwrap/rewrap the OLE2 (MS-CFB) container around a legacy `.xls` (BIFF8)
 * embedding's `Workbook` stream.
 *
 * A real `Excel.Sheet.8` object PowerPoint embeds (`Shapes.AddOLEObject`
 * against a `.xls` file, verified via COM) is a full compound binary file:
 * the BIFF8 record stream `ole-sheet-xls-biff8*.ts` reads/writes lives
 * inside its `Workbook` (or legacy `Book`) stream, alongside sibling
 * streams such as `\x05SummaryInformation` that must survive an edit
 * byte-for-byte. `readOleXlsGrid`/`writeOleXlsNumericCellEdit`/
 * `writeOleXlsStringCellEdit` all operate on the plain BIFF8 stream, so this
 * module is the seam that lets them do that against the real container
 * while a unit test can keep passing a bare BIFF8 array directly (no CFB
 * wrapper): `unwrapXlsBytes` recognises a non-CFB input and hands the bytes
 * back unchanged with no `rewrap`, matching the pre-CFB-aware behaviour.
 *
 * @module ole-sheet-xls-cfb
 */
import { parseOle2 } from './ole2-parser-read.js';
import { ENTRY_TYPE_ROOT, ENTRY_TYPE_STREAM, OLE_MAGIC } from './ole2-parser-types.js';
import { buildOle2 } from './ole2-parser-write.js';
import { readCompoundFileStream, replaceCompoundFileStream } from './ole2-stream-edit.js';

const WORKBOOK_STREAM_NAMES = ['workbook', 'book'];

export interface XlsCfbUnwrap {
	/** The plain BIFF8 record stream to read or edit. */
	workbookBytes: Uint8Array;
	/**
	 * Rebuild a flat, root-only compound file with an edited `Workbook` stream.
	 * Unsupported directory layouts return the exact original file instead of
	 * flattening storage paths and risking data loss.
	 * Undefined when the original `bytes` were not themselves a compound
	 * file (already a bare BIFF8 stream): callers use the edited bytes
	 * directly in that case.
	 */
	rewrap?: (editedWorkbookBytes: Uint8Array) => Uint8Array;
}

/** Recognise and unwrap a `.xls` payload's CFB container, when it has one. */
export function unwrapXlsBytes(bytes: Uint8Array): XlsCfbUnwrap {
	if (!isCfb(bytes)) {
		return { workbookBytes: bytes };
	}
	try {
		const buffer = bytes.buffer.slice(
			bytes.byteOffset,
			bytes.byteOffset + bytes.byteLength,
		) as ArrayBuffer;
		const ole = parseOle2(buffer);
		const workbookEntries = ole.entries.filter(
			(e) => e.type === ENTRY_TYPE_STREAM && WORKBOOK_STREAM_NAMES.includes(e.name.toLowerCase()),
		);
		const rootWorkbooks = WORKBOOK_STREAM_NAMES.map((name) => ({
			name,
			data: readCompoundFileStream(bytes, [name]),
		})).filter((item) => item.data !== undefined);
		if (rootWorkbooks.length !== 1) return { workbookBytes: bytes };
		const rootWorkbook = rootWorkbooks[0];
		const streamEntry = rootWorkbook
			? workbookEntries.find((entry) => entry.name.toLowerCase() === rootWorkbook.name)
			: undefined;
		const workbookBytes = rootWorkbook?.data;
		if (!rootWorkbook || !streamEntry || !workbookBytes) {
			return { workbookBytes: bytes };
		}
		const rootClsid = ole.entries.find((e) => e.type === ENTRY_TYPE_ROOT)?.clsid;
		const originalBytes = bytes;
		const canRebuild = hasFlatUniqueStreamLayout(ole.entries, workbookEntries.length);
		return {
			workbookBytes,
			rewrap: (editedWorkbookBytes) => {
				if (equalBytes(editedWorkbookBytes, workbookBytes)) return originalBytes;
				if (editedWorkbookBytes.length === workbookBytes.length) {
					return replaceCompoundFileStream(originalBytes, [streamEntry.name], editedWorkbookBytes);
				}
				if (!canRebuild) return originalBytes;
				const streams = new Map<string, Uint8Array>();
				for (const entry of ole.entries) {
					if (entry.type !== ENTRY_TYPE_STREAM) {
						continue;
					}
					const data =
						entry.name.toLowerCase() === streamEntry.name.toLowerCase()
							? editedWorkbookBytes
							: ole.getStream(entry.name);
					if (data) {
						streams.set(entry.name, data);
					}
				}
				return new Uint8Array(buildOle2(streams, rootClsid));
			},
		};
	} catch {
		return { workbookBytes: bytes };
	}
}

function hasFlatUniqueStreamLayout(
	entries: ReturnType<typeof parseOle2>['entries'],
	workbookCount: number,
): boolean {
	const rootEntries = entries.filter((entry) => entry.type === ENTRY_TYPE_ROOT);
	const streamEntries = entries.filter((entry) => entry.type === ENTRY_TYPE_STREAM);
	if (
		rootEntries.length !== 1 ||
		workbookCount !== 1 ||
		entries.length !== rootEntries.length + streamEntries.length
	)
		return false;
	const names = streamEntries.map((entry) => entry.name.toLowerCase());
	return new Set(names).size === names.length;
}

function equalBytes(left: Uint8Array, right: Uint8Array): boolean {
	return left.length === right.length && left.every((value, index) => value === right[index]);
}

function isCfb(bytes: Uint8Array): boolean {
	return OLE_MAGIC.every((value, index) => bytes[index] === value);
}
