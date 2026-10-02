import { describe, expect, it } from 'vitest';

import { readXlsWorkbook, XlsReadError } from '../src/legacy-excel-workbook.js';
import {
	decodeFormula,
	quoteSheetName,
	type PtgContext,
} from '../src/legacy-excel-workbook-ptg.js';
import { readXlsRecords } from '../src/legacy-excel-workbook-records.js';
import { parseSharedStrings } from '../src/legacy-excel-workbook-sst.js';
import { buildOle2 } from '../src/ole2-parser-write.js';

const u16 = (value: number): number[] => [value & 0xff, (value >> 8) & 0xff];
const u32 = (value: number): number[] => [...u16(value & 0xffff), ...u16(value >>> 16)];
const f64 = (value: number): number[] => {
	const bytes = new Uint8Array(8);
	new DataView(bytes.buffer).setFloat64(0, value, true);
	return [...bytes];
};
const ascii = (text: string): number[] => [...text].map((char) => char.charCodeAt(0));
const wide = (text: string): number[] => [...text].flatMap((char) => u16(char.charCodeAt(0)));
const record = (opcode: number, data: number[]): number[] => [
	...u16(opcode),
	...u16(data.length),
	...data,
];
const bof = (type: number, version = 0x0600): number[] =>
	record(0x0809, [...u16(version), ...u16(type), ...new Array<number>(12).fill(0)]);
const EOF = record(0x000a, []);

/** A minimal globals + one worksheet stream with the given sheet records. */
function workbookStream(
	sheetRecords: number[],
	globalsExtra: number[] = [],
	offsetDelta = 0,
): Uint8Array {
	const name = ascii('S1');
	const globalsHead = [...bof(0x0005), ...globalsExtra];
	const boundSheetLength = 4 + 8 + name.length;
	const sheetOffset = globalsHead.length + boundSheetLength + EOF.length;
	const boundSheet = record(0x0085, [
		...u32(sheetOffset + offsetDelta),
		0,
		0,
		name.length,
		0,
		...name,
	]);
	return new Uint8Array([
		...globalsHead,
		...boundSheet,
		...EOF,
		...bof(0x0010),
		...sheetRecords,
		...EOF,
	]);
}

describe('BIFF8 record framing and shared strings', () => {
	it('attaches CONTINUE records and re-reads the high-byte flag at each boundary', () => {
		// "ABCDE" starts compressed, continues wide ("中F") and the next string follows normally.
		const sst = record(0x00fc, [...u32(2), ...u32(2), ...u16(5), 0x00, ...ascii('ABC')]);
		const continued = record(0x003c, [0x01, ...wide('中F'), ...u16(2), 0x00, ...ascii('ok')]);
		const records = readXlsRecords(new Uint8Array([...sst, ...continued]));
		expect(records).toHaveLength(1);
		expect(records[0]?.continues).toHaveLength(1);
		expect(parseSharedStrings(records[0]!)).toEqual(['ABC中F', 'ok']);
	});

	it('skips rich-text runs and phonetic blocks', () => {
		const rich = [...u16(2), 0x0c, ...u16(1), ...u32(3), ...ascii('hi'), ...u32(0), 1, 2, 3];
		const sst = record(0x00fc, [...u32(2), ...u32(2), ...rich, ...u16(1), 0x00, ...ascii('z')]);
		expect(parseSharedStrings(readXlsRecords(new Uint8Array(sst))[0]!)).toEqual(['hi', 'z']);
	});
});

describe('BIFF8 cell records', () => {
	it('reads MULRK, MULBLANK, LABEL, RSTRING, BOOLERR and NUMBER cells from a bare stream', () => {
		const rkInt = (n: number) => ((n << 2) | 2) >>> 0;
		const stream = workbookStream([
			...record(0x00bd, [
				...u16(0),
				...u16(0),
				...u16(0),
				...u32(rkInt(7)),
				...u16(0),
				...u32(rkInt(-3)),
				...u16(1),
			]),
			...record(0x00be, [...u16(1), ...u16(2), ...u16(0), ...u16(0), ...u16(3)]),
			...record(0x0204, [...u16(2), ...u16(0), ...u16(0), ...u16(3), 0x01, ...wide('été')]),
			...record(0x00d6, [
				...u16(2),
				...u16(1),
				...u16(0),
				...u16(2),
				0x00,
				...ascii('rs'),
				...u16(0),
			]),
			...record(0x0205, [...u16(3), ...u16(0), ...u16(0), 0x24, 0x01]),
			...record(0x0205, [...u16(3), ...u16(1), ...u16(0), 0x01, 0x00]),
			...record(0x0203, [...u16(4), ...u16(0), ...u16(0), ...f64(Math.PI)]),
		]);
		const cells = readXlsWorkbook(stream).sheets[0]?.cells ?? [];
		expect(cells.map((cell) => [cell.row, cell.col, cell.value])).toEqual([
			[0, 0, 7],
			[0, 1, -3],
			[1, 2, null],
			[1, 3, null],
			[2, 0, 'été'],
			[2, 1, 'rs'],
			[3, 0, { error: '#NUM!' }],
			[3, 1, true],
			[4, 0, Math.PI],
		]);
	});

	it('expands SHRFMLA definitions relative to each cell and reads following STRING results', () => {
		const head = (row: number, string = false) => [
			...u16(row),
			...u16(1),
			...u16(0),
			...(string ? [0, 0, 0, 0, 0, 0, 0xff, 0xff] : f64(row)),
			...u16(0x08),
			...u32(0),
		];
		const exp = [...u16(5), 0x01, ...u16(0), ...u16(1)];
		// PtgRefN: same row, one column to the left (relative), times 2.
		const shared = [0x2c, ...u16(0), ...u16(0xc0ff), 0x1e, ...u16(2), 0x05];
		const stream = workbookStream([
			...record(0x0006, [...head(0, true), ...exp]),
			...record(0x04bc, [...u16(0), ...u16(2), 1, 1, 0, 3, ...u16(shared.length), ...shared]),
			...record(0x0207, [...u16(3), 0x00, ...ascii('txt')]),
			...record(0x0006, [...head(1), ...exp]),
			...record(0x0006, [...head(2), ...exp]),
		]);
		const cells = readXlsWorkbook(stream).sheets[0]?.cells ?? [];
		expect(cells.map((cell) => [cell.formula, cell.value, cell.sharedFormula])).toEqual([
			['A1*2', 'txt', true],
			['A2*2', 1, true],
			['A3*2', 2, true],
		]);
	});

	it('flags formulas whose tokens cannot be decoded', () => {
		const formula = [...u16(0), ...u16(0), ...u16(0), ...f64(1), ...u16(0), ...u32(0)];
		const stream = workbookStream([...record(0x0006, [...formula, ...u16(1), 0x18])]);
		const cell = readXlsWorkbook(stream).sheets[0]?.cells[0];
		expect(cell).toMatchObject({ value: 1, formulaUndecoded: true });
		expect(cell?.formula).toBeUndefined();
	});

	it('skips embedded chart substreams and reports them', () => {
		const stream = workbookStream([
			...bof(0x0020),
			...record(0x0203, [...u16(9), ...u16(9), ...u16(0), ...f64(1)]),
			...EOF,
		]);
		const sheet = readXlsWorkbook(stream).sheets[0];
		expect(sheet?.cells).toEqual([]);
		expect(sheet?.unsupported).toContain('charts');
	});
});

describe('BIFF8 formula tokens', () => {
	const context: PtgContext = {
		sheetPrefix: (ixti) => (ixti === 0 ? 'Data' : ixti === 1 ? "'My Sheet'" : undefined),
		name: (index) => (index === 1 ? 'Rate' : undefined),
		externName: () => undefined,
	};
	const decode = (tokens: number[], extra: number[] = [], at = { row: 4, col: 2 }) =>
		decodeFormula(new Uint8Array(tokens), new Uint8Array(extra), context, at);
	const ref = (row: number, col: number, rowRel = true, colRel = true) => [
		...u16(row),
		...u16(col | (colRel ? 0x4000 : 0) | (rowRel ? 0x8000 : 0)),
	];

	it('decodes constants, operators and parentheses', () => {
		expect(decode([0x1e, ...u16(2), 0x1f, ...f64(1.5), 0x03, 0x15, 0x1e, ...u16(3), 0x05])).toBe(
			'(2+1.5)*3',
		);
		expect(
			decode([0x17, 2, 0, ...ascii('a"'), 0x1d, 1, 0x1c, 0x07, 0x16, 0x42, 4, ...u16(1)]),
		).toBe('IF("a""",TRUE,#DIV/0!,)');
		expect(decode([0x1e, ...u16(5), 0x13, 0x14])).toBe('-5%');
	});

	it('decodes absolute, relative, shared-relative and 3-D references', () => {
		expect(decode([0x24, ...ref(0, 0, false, false)])).toBe('$A$1');
		expect(decode([0x25, ...u16(0), ...u16(65535), ...u16(0x4001), ...u16(0x4001)])).toBe('B:B');
		// RefN offsets are relative to the formula cell (row 4, col 2): -1 row, +1 col.
		expect(decode([0x2c, ...u16(0xffff), ...u16(0xc001)])).toBe('D4');
		expect(decode([0x3a, ...u16(1), ...ref(2, 3)])).toBe("'My Sheet'!D3");
		expect(decode([0x3b, ...u16(0), ...u16(0), ...u16(1), ...u16(0), ...u16(1)])).toBe(
			'Data!$A$1:$B$2',
		);
		expect(decode([0x2a, 0, 0, 0, 0])).toBe('#REF!');
		expect(decode([0x23, ...u32(1)])).toBe('Rate');
	});

	it('handles PtgAttr sum, choose and whitespace, and the PtgMem* wrappers', () => {
		expect(
			decode([0x25, ...u16(0), ...u16(2), ...u16(0xc000), ...u16(0xc000), 0x19, 0x10, 0, 0]),
		).toBe('SUM(A1:A3)');
		const choose = [
			0x1e,
			...u16(1),
			0x19,
			0x04,
			...u16(1),
			0,
			0,
			0,
			0,
			0x1e,
			...u16(9),
			0x42,
			2,
			...u16(100),
		];
		expect(decode(choose)).toBe('CHOOSE(1,9)');
		expect(decode([0x1e, ...u16(1), 0x1e, ...u16(2), 0x19, 0x40, 0, 2, 0x03])).toBe('1  +2');
		expect(decode([0x29, ...u16(5), 0x24, ...ref(0, 0), 0x41, ...u16(24)])).toBe('ABS(A1)');
	});

	it('returns undefined for unknown tokens and functions', () => {
		expect(decode([0x18, 0x01])).toBeUndefined();
		expect(decode([0x21, ...u16(4000)])).toBeUndefined();
		expect(decode([0x3a, ...u16(7), ...ref(0, 0)])).toBeUndefined();
	});

	it('quotes sheet names the way Excel does', () => {
		expect(quoteSheetName('Data')).toBe('Data');
		expect(quoteSheetName('My Sheet')).toBe("'My Sheet'");
		expect(quoteSheetName("Bob's")).toBe("'Bob''s'");
		expect(quoteSheetName('A1')).toBe("'A1'");
		expect(quoteSheetName('2024')).toBe("'2024'");
	});
});

describe('readXlsWorkbook errors', () => {
	it('rejects BIFF5 and FILEPASS streams with typed codes', () => {
		const biff5 = new Uint8Array([...bof(0x0005, 0x0500), ...EOF]);
		expect(() => readXlsWorkbook(biff5)).toThrow(
			expect.objectContaining({ code: 'unsupported-version' }),
		);
		const encrypted = workbookStream([], record(0x002f, [1, 0, 1, 0, 1, 0]));
		expect(() => readXlsWorkbook(encrypted)).toThrow(
			expect.objectContaining({ code: 'encrypted' }),
		);
	});

	it('rejects compound files without a Workbook stream and BIFF5 Book streams', () => {
		const other = buildOle2(new Map([['Other', new Uint8Array(8)]]));
		expect(() => readXlsWorkbook(new Uint8Array(other))).toThrow(XlsReadError);
		const book = buildOle2(new Map([['Book', new Uint8Array([...bof(0x0005, 0x0500), ...EOF])]]));
		expect(() => readXlsWorkbook(new Uint8Array(book))).toThrow(/BIFF5/);
	});

	it('falls back to file order when a BOUNDSHEET offset is wrong', () => {
		const stream = workbookStream([...record(0x0203, [0, 0, 0, 0, 0, 0, ...f64(2)])], [], 999);
		expect(readXlsWorkbook(stream).sheets[0]?.cells[0]?.value).toBe(2);
	});

	it('accepts a compound file wrapping a minimal stream', () => {
		const cfb = buildOle2(new Map([['Workbook', workbookStream([])]]));
		expect(readXlsWorkbook(new Uint8Array(cfb)).sheets.map((sheet) => sheet.name)).toEqual(['S1']);
	});
});
