/**
 * Full read-only BIFF8 workbook reader for Excel 97-2003 `.xls` files.
 *
 * `readXlsWorkbook` unwraps the compound file's `Workbook` stream, parses the
 * globals substream (sheets, shared strings, fonts, formats, XFs, palette,
 * names, external references) and every worksheet substream (cells with
 * cached formula results and decoded formula text, merges, column and row
 * sizes, view and panes, hyperlinks, comments). Chart, macro and VBA module
 * sheets are listed with their kind but no cells. Content the reader does not
 * model is reported in `unsupported` lists rather than silently dropped.
 *
 * Encrypted workbooks (`FILEPASS`), pre-BIFF8 files and unreadable input throw
 * {@link XlsReadError} with a `code`.
 *
 * Reference: [MS-XLS] Excel Binary File Format.
 * @module legacy-excel-workbook
 */
import { unwrapXlsBytes } from './legacy-excel-cfb.js';
import { parseGlobals } from './legacy-excel-workbook-globals.js';
import { readXlsRecords, type XlsRecord } from './legacy-excel-workbook-records.js';
import { parseSheet } from './legacy-excel-workbook-sheet.js';
import { parseOle2 } from './ole2-parser-read.js';
import { OLE_MAGIC } from './ole2-parser-types.js';
import { XlsReadError, type XlsSheet, type XlsWorkbook } from './legacy-excel-workbook-types.js';

export * from './legacy-excel-workbook-types.js';
export { XLS_BUILTIN_FORMATS, DEFAULT_XLS_PALETTE } from './legacy-excel-workbook-styles.js';

function compoundStreamNames(bytes: Uint8Array): string[] {
	if (!OLE_MAGIC.every((value, index) => bytes[index] === value)) return [];
	try {
		const buffer = bytes.buffer.slice(
			bytes.byteOffset,
			bytes.byteOffset + bytes.byteLength,
		) as ArrayBuffer;
		return parseOle2(buffer).entries.map((entry) => entry.name);
	} catch {
		throw new XlsReadError('corrupt', 'The file is not a readable OLE2 compound file.');
	}
}

/** Record indices of the `BOF` of every substream after the globals, in file order. */
function topLevelSubstreams(records: XlsRecord[]): number[] {
	const starts: number[] = [];
	let depth = 0;
	records.forEach((record, index) => {
		if (record.opcode === 0x0809) {
			if (depth === 0 && index > 0) starts.push(index);
			depth++;
		} else if (record.opcode === 0x000a) depth = Math.max(0, depth - 1);
	});
	return starts;
}

/**
 * Read a BIFF8 `.xls` workbook. Accepts the whole compound file or a bare
 * `Workbook` stream.
 */
export function readXlsWorkbook(bytes: Uint8Array): XlsWorkbook {
	const streams = compoundStreamNames(bytes);
	if (
		streams.some((name) => name.toLowerCase() === 'book') &&
		!streams.some((name) => name.toLowerCase() === 'workbook')
	)
		throw new XlsReadError(
			'unsupported-version',
			'This is an Excel 5.0/95 (BIFF5) workbook; only Excel 97-2003 (BIFF8) files are supported.',
		);
	if (streams.some((name) => name === 'EncryptedPackage'))
		throw new XlsReadError('encrypted', 'The workbook is password protected (encrypted).');
	const stream = unwrapXlsBytes(bytes).workbookBytes;
	if (streams.length > 0 && stream === bytes)
		throw new XlsReadError('corrupt', 'The compound file has no Workbook stream.');
	const records = readXlsRecords(stream);
	if (records.length === 0) throw new XlsReadError('corrupt', 'The workbook stream is empty.');
	const globals = parseGlobals(records);

	const unsupported = [...globals.unsupported];
	if (
		streams.some((name) => name === '_VBA_PROJECT_CUR' || name === 'VBA') &&
		!unsupported.includes('VBA macros')
	)
		unsupported.push('VBA macros');
	if (streams.some((name) => name === '_SX_DB_CUR') && !unsupported.includes('pivot caches'))
		unsupported.push('pivot caches');

	const byOffset = new Map(records.map((record, index) => [record.offset, index]));
	const substreams = topLevelSubstreams(records);
	const sheets: XlsSheet[] = globals.sheets.map((bound, position) => {
		// Fall back to file order when a writer recorded a wrong BOUNDSHEET offset.
		const start = byOffset.get(bound.offset) ?? substreams[position];
		if (start === undefined || bound.kind !== 'worksheet') {
			const empty = parseSheet(bound, [], 0, globals);
			if (start === undefined && bound.kind === 'worksheet')
				empty.unsupported.push('sheet data (substream not found)');
			return empty;
		}
		return parseSheet(bound, records, start, globals);
	});

	return {
		biffVersion: 8,
		date1904: globals.date1904,
		sheets,
		fonts: globals.fonts,
		xfs: globals.xfs,
		numberFormats: globals.numberFormats,
		palette: globals.palette,
		names: globals.names,
		activeSheet: globals.activeSheet,
		...(globals.codepage === undefined ? {} : { codepage: globals.codepage }),
		structureLocked: globals.structureLocked,
		unsupported,
	};
}
