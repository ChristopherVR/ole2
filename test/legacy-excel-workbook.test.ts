import { readFileSync } from 'node:fs';
import { describe, expect, it } from 'vitest';

import {
	readXlsWorkbook,
	XlsReadError,
	type XlsCell,
	type XlsSheet,
	type XlsWorkbook,
} from '../src/legacy-excel-workbook.js';

// Fixtures are real Excel 16 saves (FileFormat 56); see fixtures/xls/generate-xls-fixtures.ps1.
function fixture(name: string): Uint8Array {
	return new Uint8Array(readFileSync(new URL(`./fixtures/xls/${name}`, import.meta.url)));
}

function cellAt(sheet: XlsSheet | undefined, row: number, col: number): XlsCell | undefined {
	return sheet?.cells.find((cell) => cell.row === row && cell.col === col);
}

describe('readXlsWorkbook: values and sheets', () => {
	const workbook = readXlsWorkbook(fixture('workbook-features.xls'));
	const data = workbook.sheets[0];

	it('lists every sheet with its visibility', () => {
		expect(workbook.biffVersion).toBe(8);
		expect(workbook.sheets.map((sheet) => [sheet.name, sheet.state, sheet.kind])).toEqual([
			['Data', 'visible', 'worksheet'],
			['My Formulas', 'visible', 'worksheet'],
			['Hidden', 'hidden', 'worksheet'],
			['Secret', 'veryHidden', 'worksheet'],
		]);
		expect(workbook.activeSheet).toBe(0);
		expect(workbook.date1904).toBe(false);
		expect(workbook.unsupported).toEqual([]);
	});

	it('reads strings, unicode, numbers, booleans and errors', () => {
		expect(cellAt(data, 0, 0)?.value).toBe('Name');
		expect(cellAt(data, 1, 1)?.value).toBe(1.25);
		expect(cellAt(data, 2, 2)?.value).toBe(24);
		expect(cellAt(data, 4, 0)?.value).toBe('日本語 été Ж');
		expect(cellAt(data, 4, 1)?.value).toBe(-1234567.891);
		expect(cellAt(data, 4, 2)?.value).toBe(123456789);
		expect(cellAt(data, 1, 3)?.value).toBe(true);
		expect(cellAt(data, 2, 3)?.value).toBe(false);
		expect(cellAt(data, 1, 4)).toMatchObject({ value: { error: '#DIV/0!' }, formula: '1/0' });
		expect(cellAt(data, 2, 4)).toMatchObject({ value: { error: '#N/A' }, formula: 'NA()' });
	});

	it('stitches shared strings across CONTINUE records, narrow and wide', () => {
		expect(cellAt(data, 7, 0)?.value).toBe('0123456789'.repeat(1000));
		expect(cellAt(data, 8, 0)?.value).toBe('ÅΩ中x'.repeat(3000));
		for (const i of [1, 300, 600]) expect(cellAt(data, 9 + i, 0)?.value).toBe(`Item number ${i}`);
	});

	it('resolves number formats for dates, times and percentages', () => {
		const format = (row: number, col: number) =>
			workbook.xfs[cellAt(data, row, col)?.xf ?? 0]?.numFmt;
		expect(cellAt(data, 1, 5)?.value).toBe(45292);
		// Excel maps yyyy-mm-dd onto the locale short-date id 14.
		expect(workbook.xfs[cellAt(data, 1, 5)?.xf ?? 0]?.numFmtId).toBe(14);
		expect(format(2, 5)).toBe('h:mm AM/PM');
		expect(format(1, 6)).toBe('0.00%');
		expect(workbook.numberFormats.get(44)).toContain('#,##0.00');
	});

	it('reads defined names including built-in names', () => {
		expect(workbook.names).toContainEqual(
			expect.objectContaining({ name: 'TaxRate', formula: 'Data!$B$2', builtin: false }),
		);
		expect(workbook.names).toContainEqual(
			expect.objectContaining({ name: 'Print_Area', formula: 'Data!$A$1:$C$5', localSheet: 0 }),
		);
	});
});

describe('readXlsWorkbook: formulas', () => {
	const workbook = readXlsWorkbook(fixture('workbook-features.xls'));
	const sheet = workbook.sheets[1];
	const formula = (row: number, col = 0) => cellAt(sheet, row, col)?.formula;

	it('decodes common functions, operators and references', () => {
		expect(formula(0)).toBe('SUM(Data!B2:B4)');
		expect(formula(1)).toBe('IF(Data!B2>1,"big","small")');
		expect(formula(2)).toBe('VLOOKUP("Banana",Data!A2:C4,3,FALSE)');
		expect(formula(3)).toBe('Data!B2*Data!C2+1');
		expect(formula(4)).toBe('CONCATENATE(Data!A2," & ",Data!A3)');
		expect(formula(5)).toBe('ROUND(AVERAGE(Data!B2:B4),2)');
		expect(formula(6)).toBe('A1-A4/2');
		expect(formula(7)).toBe('COUNTIF(Data!C2:C4,">8")');
		expect(formula(8)).toBe('$A$1+A$2&"x"');
		expect(formula(9)).toBe('MAX(1,2,3)-MIN(A1:A4)');
		expect(formula(12)).toBe('-A1^2%');
		expect(formula(14)).toBe('IF(A1>0,SUM(A1:A2),"")');
		expect(formula(17)).toBe('CHOOSE(2,"x","y","z")');
		expect(formula(18)).toBe('SUM(Data!B:B)');
		expect(formula(19)).toBe('Hidden!A1');
		expect(formula(21)).toBe('SUMPRODUCT(Data!B2:B4,Data!C2:C4)');
	});

	it('decodes add-in and Excel 2007 functions, whitespace, array constants and 3-D ranges', () => {
		expect(formula(22)).toBe('IFERROR(1/0,"err")');
		expect(formula(23)).toBe('SUMIFS(Data!B2:B4,Data!C2:C4,">8")');
		expect(formula(24)).toBe('EOMONTH(Data!F2,1)');
		expect(formula(25)).toBe('A1 + A4');
		expect(formula(26)).toBe('SUM({1,2;3,4})');
		expect(formula(27)).toBe('SUM(Data:Hidden!Z1:Z2)');
		expect(formula(28)).toBe(`Secret!A1&" "&'My Formulas'!A2`);
		expect(formula(0, 3)).toBe('TaxRate*100');
	});

	it('keeps cached results of every type, including STRING records', () => {
		expect(cellAt(sheet, 0, 0)?.value).toBe(5.5);
		expect(cellAt(sheet, 1, 0)?.value).toBe('big');
		expect(cellAt(sheet, 4, 0)?.value).toBe('Apple & Banana');
		expect(cellAt(sheet, 10, 0)?.value).toBe(true);
		expect(cellAt(sheet, 15, 0)?.value).toEqual({ error: '#DIV/0!' });
		expect(cellAt(sheet, 28, 0)?.value).toBe('shh big');
	});

	it('reads range-filled formulas per cell and anchors array formulas', () => {
		// Excel 2007+ writes one FORMULA per cell (no SHRFMLA) when saving .xls.
		expect([0, 1, 2, 5].map((row) => formula(row, 1))).toEqual([
			'Data!B2*2',
			'Data!B3*2',
			'Data!B4*2',
			'Data!B7*2',
		]);
		expect(cellAt(sheet, 0, 2)).toMatchObject({
			formula: 'Data!B2:B4*10',
			value: 12.5,
			arrayRange: { firstRow: 0, lastRow: 2, firstCol: 2, lastCol: 2 },
		});
		expect(cellAt(sheet, 1, 2)).toMatchObject({ value: 5 });
		expect(cellAt(sheet, 1, 2)?.formula).toBeUndefined();
	});
});

describe('readXlsWorkbook: formatting and sheet layout', () => {
	const workbook: XlsWorkbook = readXlsWorkbook(fixture('workbook-styles.xls'));
	const sheet = workbook.sheets[0];
	const xfOf = (row: number, col: number) => workbook.xfs[cellAt(sheet, row, col)?.xf ?? 0];
	const fontOf = (row: number, col: number) => workbook.fonts[xfOf(row, col)?.font ?? 0];

	it('resolves fonts and colours', () => {
		expect(fontOf(0, 0)).toMatchObject({ bold: true, color: { rgb: 'FF0000' } });
		expect(fontOf(1, 0)).toMatchObject({ name: 'Arial', size: 14, italic: true });
		expect(fontOf(3, 0)).toMatchObject({ underline: 'single', strike: true });
		expect(xfOf(6, 0)?.fontColor).toEqual({ theme: 10 });
	});

	it('resolves fills, borders, alignment and number formats', () => {
		expect(xfOf(2, 0)?.fill).toMatchObject({ pattern: 'solid', fg: { rgb: 'FFFF00' } });
		expect(xfOf(5, 1)?.fill).toMatchObject({
			pattern: 'darkGrid',
			fg: { rgb: 'FF0000' },
			bg: { rgb: 'CCFFFF' },
		});
		expect(xfOf(1, 1)?.border).toMatchObject({
			left: { style: 'thin' },
			top: { style: 'medium' },
			bottom: { style: 'double', color: { rgb: '0000FF' } },
			right: { style: 'dashed' },
		});
		expect(xfOf(4, 1)?.alignment).toMatchObject({
			horizontal: 'center',
			vertical: 'top',
			wrap: true,
		});
		expect(xfOf(0, 2)?.alignment.rotation).toBe(45);
		expect(xfOf(1, 2)?.alignment.indent).toBe(2);
		expect(xfOf(2, 1)?.numFmt).toBe('#,##0.00');
		expect(xfOf(3, 1)?.numFmt).toBe('"USD"\\ #,##0.000;[Red]\\-#,##0.000');
	});

	it('reads merges, column widths, row heights and hidden rows/columns', () => {
		expect(sheet?.merges).toEqual([
			{ firstRow: 0, lastRow: 1, firstCol: 3, lastCol: 5 },
			{ firstRow: 3, lastRow: 3, firstCol: 3, lastCol: 4 },
		]);
		expect(sheet?.columns.find((col) => col.firstCol === 0)?.width).toBeCloseTo(25.63, 1);
		expect(sheet?.columns.find((col) => col.firstCol === 6)?.hidden).toBe(true);
		expect(sheet?.rows.find((row) => row.row === 1)).toMatchObject({
			height: 30,
			customHeight: true,
		});
		expect(sheet?.rows.find((row) => row.row === 8)?.hidden).toBe(true);
		expect(sheet?.defaultRowHeight).toBe(14.5);
	});

	it('reads the view: frozen panes, gridlines, zoom, selection and tab colour', () => {
		expect(sheet?.view).toMatchObject({
			showGridLines: false,
			zoom: 85,
			freeze: { rows: 2, cols: 1 },
			activeCell: { row: 2, col: 1 },
		});
		expect(workbook.sheets[1]?.tabColor?.rgb).toBe('FF0000');
	});

	it('reads hyperlinks and comments', () => {
		expect(sheet?.hyperlinks).toEqual([
			{
				range: { firstRow: 6, lastRow: 6, firstCol: 0, lastCol: 0 },
				display: 'Example site',
				target: 'https://example.com/path?q=1',
				tooltip: 'Example tip',
			},
			{
				range: { firstRow: 7, lastRow: 7, firstCol: 0, lastCol: 0 },
				display: 'Jump to Other',
				location: 'Other!B3',
			},
		]);
		expect(sheet?.comments).toEqual([
			{ row: 5, col: 0, author: 'Author', text: 'Reviewed by QA', visible: false },
		]);
	});
});

describe('readXlsWorkbook: workbook modes and errors', () => {
	it('reads the 1904 date system', () => {
		const workbook = readXlsWorkbook(fixture('workbook-1904.xls'));
		expect(workbook.date1904).toBe(true);
		expect(cellAt(workbook.sheets[0], 1, 0)?.value).toBe(43830);
	});

	it('throws a typed error for an encrypted workbook', () => {
		expect(() => readXlsWorkbook(fixture('workbook-encrypted.xls'))).toThrow(XlsReadError);
		try {
			readXlsWorkbook(fixture('workbook-encrypted.xls'));
		} catch (error) {
			expect((error as XlsReadError).code).toBe('encrypted');
		}
	});

	it('rejects non-workbook input as corrupt', () => {
		expect(() => readXlsWorkbook(new Uint8Array(16))).toThrow(XlsReadError);
		expect(() => readXlsWorkbook(new Uint8Array())).toThrow(/empty/);
	});
});
