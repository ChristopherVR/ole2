/**
 * The BIFF8 workbook globals substream: sheet directory (`BOUNDSHEET`),
 * shared strings, formatting tables, `DATEMODE`, `WINDOW1`, `PROTECT`,
 * `CODEPAGE`, defined names (`NAME`) and the external reference tables
 * (`SUPBOOK`, `EXTERNNAME`, `EXTERNSHEET`) that 3-D formula tokens index.
 *
 * @module legacy-excel-workbook-globals
 */
import { decodeFormula, quoteSheetName, type PtgContext } from './legacy-excel-workbook-ptg.js';
import { ByteReader, joinedData, type XlsRecord } from './legacy-excel-workbook-records.js';
import { parseSharedStrings } from './legacy-excel-workbook-sst.js';
import {
	applyXfExt,
	DEFAULT_XLS_PALETTE,
	parseFont,
	parseFormat,
	parsePalette,
	parseXf,
	XLS_BUILTIN_FORMATS,
} from './legacy-excel-workbook-styles.js';
import {
	XlsReadError,
	type XlsFont,
	type XlsName,
	type XlsSheetKind,
	type XlsSheetState,
	type XlsXf,
} from './legacy-excel-workbook-types.js';

export const RT = {
	BOF: 0x0809,
	EOF: 0x000a,
	FILEPASS: 0x002f,
	CODEPAGE: 0x0042,
	DATEMODE: 0x0022,
	PROTECT: 0x0012,
	WINDOW1: 0x003d,
	FONT: 0x0031,
	FORMAT: 0x041e,
	XF: 0x00e0,
	XFEXT: 0x087d,
	PALETTE: 0x0092,
	BOUNDSHEET: 0x0085,
	SUPBOOK: 0x01ae,
	EXTERNNAME: 0x0023,
	EXTERNSHEET: 0x0017,
	NAME: 0x0018,
	SST: 0x00fc,
	OBPROJ: 0x00d3,
	SXSTREAMID: 0x00d5,
	MSODRAWINGGROUP: 0x00eb,
} as const;

export interface XlsBoundSheet {
	name: string;
	state: XlsSheetState;
	kind: XlsSheetKind;
	/** Stream offset of the sheet's `BOF`. */
	offset: number;
}

interface SupBook {
	kind: 'self' | 'addin' | 'external';
	sheets: string[];
	path: string;
	externNames: string[];
}

export interface XlsGlobals {
	sheets: XlsBoundSheet[];
	sst: string[];
	fonts: XlsFont[];
	xfs: XlsXf[];
	numberFormats: Map<number, string>;
	palette: string[];
	names: XlsName[];
	date1904: boolean;
	activeSheet: number;
	codepage?: number;
	structureLocked: boolean;
	unsupported: string[];
	context: PtgContext;
}

const STATES: readonly XlsSheetState[] = ['visible', 'hidden', 'veryHidden'];
const KINDS: Readonly<Record<number, XlsSheetKind>> = {
	0: 'worksheet',
	1: 'macro',
	2: 'chart',
	6: 'vba',
};

const BUILTIN_NAMES = [
	'Consolidate_Area',
	'Auto_Open',
	'Auto_Close',
	'Extract',
	'Database',
	'Criteria',
	'Print_Area',
	'Print_Titles',
	'Recorder',
	'Data_Form',
	'Auto_Activate',
	'Auto_Deactivate',
	'Sheet_Title',
	'_FilterDatabase',
];

interface RawName {
	name: XlsName;
	rgce: Uint8Array;
	extra: Uint8Array;
}

function parseName(data: Uint8Array): RawName {
	const r = new ByteReader(data);
	const flags = r.u16();
	r.skip(1);
	const length = r.u8();
	const cce = r.u16();
	r.skip(2);
	const itab = r.u16();
	r.skip(4);
	const wide = (r.u8() & 1) !== 0;
	const builtin = (flags & 0x20) !== 0;
	let text = r.chars(length, wide);
	if (builtin && text.length === 1) text = BUILTIN_NAMES[text.charCodeAt(0)] ?? text;
	const rgce = data.subarray(r.pos, r.pos + cce);
	const name: XlsName = {
		name: text,
		hidden: (flags & 0x01) !== 0,
		builtin,
		isFunction: (flags & 0x02) !== 0,
	};
	if (itab > 0) name.localSheet = itab - 1;
	return { name, rgce, extra: data.subarray(r.pos + cce) };
}

function parseSupBook(data: Uint8Array): SupBook {
	const r = new ByteReader(data);
	const count = r.u16();
	const marker = r.u16();
	if (marker === 0x0401) return { kind: 'self', sheets: [], path: '', externNames: [] };
	if (marker === 0x3a01) return { kind: 'addin', sheets: [], path: '', externNames: [] };
	r.pos = 2;
	const rawPath = r.xlString();
	const sheets: string[] = [];
	for (let i = 0; i < count && r.remaining > 0; i++) sheets.push(r.xlString());
	// VirtualPath encodes path separators as control characters; keep the file name.
	const path = rawPath.split(/[\u0001-\u0008\\/]/).pop() ?? rawPath;
	return { kind: 'external', sheets, path, externNames: [] };
}

function sheetRange(first: string, last: string): string {
	if (first === last) return quoteSheetName(first);
	const plain = quoteSheetName(first) === first && quoteSheetName(last) === last;
	return plain ? `${first}:${last}` : `'${`${first}:${last}`.replace(/'/g, "''")}'`;
}

/** Parse the globals substream records (from its `BOF` to its `EOF`). */
export function parseGlobals(records: XlsRecord[]): XlsGlobals {
	const first = records[0];
	if (!first || first.opcode !== RT.BOF) throw new XlsReadError('corrupt', 'Missing BOF record.');
	const version = first.data[0]! | (first.data[1]! << 8);
	if (version !== 0x0600)
		throw new XlsReadError(
			'unsupported-version',
			'Only Excel 97-2003 (BIFF8) workbooks are supported; this file uses an older Excel format.',
		);
	const end = records.findIndex((record) => record.opcode === RT.EOF);
	const globals = end >= 0 ? records.slice(0, end) : records;
	if (globals.some((record) => record.opcode === RT.FILEPASS))
		throw new XlsReadError('encrypted', 'The workbook is password protected (encrypted).');

	const paletteRecord = globals.find((record) => record.opcode === RT.PALETTE);
	const palette = paletteRecord ? parsePalette(paletteRecord.data) : [...DEFAULT_XLS_PALETTE];
	const numberFormats = new Map<number, string>();
	for (const record of globals) {
		if (record.opcode !== RT.FORMAT) continue;
		const format = parseFormat(joinedData(record));
		numberFormats.set(format.id, format.code);
	}
	const formatCode = (id: number): string =>
		numberFormats.get(id) ?? XLS_BUILTIN_FORMATS[id] ?? 'General';

	const out: XlsGlobals = {
		sheets: [],
		sst: [],
		fonts: [],
		xfs: [],
		numberFormats,
		palette,
		names: [],
		date1904: false,
		activeSheet: 0,
		structureLocked: false,
		unsupported: [],
		context: { sheetPrefix: () => undefined, name: () => undefined, externName: () => undefined },
	};
	const supbooks: SupBook[] = [];
	const xtis: { book: number; first: number; last: number }[] = [];
	const rawNames: RawName[] = [];
	for (const record of globals) {
		const r = new ByteReader(record.data);
		switch (record.opcode) {
			case RT.FONT:
				out.fonts.push(parseFont(record.data, palette));
				break;
			case RT.XF:
				out.xfs.push(parseXf(record.data, palette, formatCode));
				break;
			case RT.XFEXT:
				applyXfExt(record.data, out.xfs, palette);
				break;
			case RT.DATEMODE:
				out.date1904 = r.u16() === 1;
				break;
			case RT.CODEPAGE:
				out.codepage = r.u16();
				break;
			case RT.PROTECT:
				out.structureLocked = r.u16() === 1;
				break;
			case RT.WINDOW1:
				r.skip(10);
				out.activeSheet = r.u16();
				break;
			case RT.BOUNDSHEET: {
				const offset = r.u32();
				const state = STATES[r.u8() & 0x03] ?? 'visible';
				const kind = KINDS[r.u8()] ?? 'unknown';
				out.sheets.push({ offset, state, kind, name: r.shortXlString() });
				break;
			}
			case RT.SST:
				out.sst = parseSharedStrings(record);
				break;
			case RT.SUPBOOK:
				supbooks.push(parseSupBook(joinedData(record)));
				break;
			case RT.EXTERNNAME: {
				const book = supbooks[supbooks.length - 1];
				const nr = new ByteReader(joinedData(record), 6);
				book?.externNames.push(nr.shortXlString());
				break;
			}
			case RT.EXTERNSHEET: {
				const xr = new ByteReader(joinedData(record));
				const count = xr.u16();
				for (let i = 0; i < count; i++)
					xtis.push({ book: xr.u16(), first: xr.i16(), last: xr.i16() });
				break;
			}
			case RT.NAME:
				rawNames.push(parseName(joinedData(record)));
				break;
			case RT.OBPROJ:
				out.unsupported.push('VBA macros');
				break;
			case RT.SXSTREAMID:
				if (!out.unsupported.includes('pivot caches')) out.unsupported.push('pivot caches');
				break;
		}
	}

	const sheetNames = out.sheets.map((sheet) => sheet.name);
	out.context = {
		sheetPrefix(ixti) {
			const xti = xtis[ixti];
			const book = xti ? supbooks[xti.book] : undefined;
			if (!xti || !book) return undefined;
			if (xti.first === -1) return '#REF';
			if (xti.first < 0) return undefined;
			if (book.kind === 'self') {
				const a = sheetNames[xti.first];
				const b = sheetNames[xti.last] ?? a;
				return a === undefined || b === undefined ? undefined : sheetRange(a, b);
			}
			if (book.kind === 'external') {
				const sheet = book.sheets[xti.first] ?? '';
				return `'[${book.path}]${sheet.replace(/'/g, "''")}'`;
			}
			return undefined;
		},
		name(index) {
			const raw = rawNames[index - 1];
			return raw?.name.name;
		},
		externName(ixti, index) {
			const xti = xtis[ixti];
			const book = xti ? supbooks[xti.book] : undefined;
			if (!book) return undefined;
			if (book.kind === 'self') return this.name(index);
			return book.externNames[index - 1];
		},
	};
	for (const raw of rawNames) {
		const formula = decodeFormula(raw.rgce, raw.extra, out.context, { row: 0, col: 0 });
		if (formula !== undefined) raw.name.formula = formula;
		out.names.push(raw.name);
	}
	return out;
}
