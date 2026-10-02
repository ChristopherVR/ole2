/**
 * Text helpers for BIFF8 formula decoding: error literals, A1 cell and area
 * text (absolute/relative markers, whole rows and columns), relative
 * (`PtgRefN`/`PtgAreaN`) location resolution, sheet-name quoting and
 * `PtgArray` constants.
 *
 * @module legacy-excel-workbook-ptg-text
 */
import type { ByteReader } from './legacy-excel-workbook-records.js';
import type { XlsErrorCode } from './legacy-excel-workbook-types.js';

export interface PtgContext {
	/** The quoted sheet prefix (without `!`) an `XTI` index points at, `#REF` for deleted sheets. */
	sheetPrefix(ixti: number): string | undefined;
	/** The defined name at a one-based `NAME` index. */
	name(index: number): string | undefined;
	/** The external or add-in name at a one-based index within the `XTI`'s supporting book. */
	externName(ixti: number, index: number): string | undefined;
}

export interface PtgOptions {
	/** The cell the formula belongs to (base for relative `PtgRefN`/`PtgAreaN`). */
	row: number;
	col: number;
	/** Shared formulas store `PtgRef3d`/`PtgArea3d` locations relative to the cell too. */
	shared?: boolean;
}

export const BIFF_ERRORS: Readonly<Record<number, XlsErrorCode>> = {
	0x00: '#NULL!',
	0x07: '#DIV/0!',
	0x0f: '#VALUE!',
	0x17: '#REF!',
	0x1d: '#NAME?',
	0x24: '#NUM!',
	0x2a: '#N/A',
	0x2b: '#GETTING_DATA',
};

export const BINARY: Readonly<Record<number, string>> = {
	0x03: '+',
	0x04: '-',
	0x05: '*',
	0x06: '/',
	0x07: '^',
	0x08: '&',
	0x09: '<',
	0x0a: '<=',
	0x0b: '=',
	0x0c: '>=',
	0x0d: '>',
	0x0e: '<>',
	0x0f: ' ',
	0x10: ',',
	0x11: ':',
};

/** Column letters for a zero-based column. */
export function columnName(col: number): string {
	let name = '';
	for (let n = col + 1; n > 0; n = Math.floor((n - 1) / 26))
		name = String.fromCharCode(65 + ((n - 1) % 26)) + name;
	return name;
}

/** Quote a sheet name for a formula when Excel would (spaces, punctuation, cell-like names). */
export function quoteSheetName(name: string): string {
	const plain =
		/^[A-Za-z_][A-Za-z0-9_.]*$/.test(name) &&
		!/^[A-Za-z]{1,3}[0-9]+$/.test(name) &&
		!/^[Rr][0-9]*([Cc][0-9]*)?$/.test(name) &&
		!/^(TRUE|FALSE)$/i.test(name);
	return plain ? name : `'${name.replace(/'/g, "''")}'`;
}

export function formatNumber(value: number): string {
	return String(value).replace('e', 'E');
}

export function stringLiteral(text: string): string {
	return `"${text.replace(/"/g, '""')}"`;
}

export interface Loc {
	row: number;
	col: number;
	rowRel: boolean;
	colRel: boolean;
}

export function cellText(loc: Loc): string {
	return `${loc.colRel ? '' : '$'}${columnName(loc.col)}${loc.rowRel ? '' : '$'}${loc.row + 1}`;
}

export function areaText(first: Loc, last: Loc): string {
	if (first.row === 0 && last.row === 0xffff)
		return `${first.colRel ? '' : '$'}${columnName(first.col)}:${last.colRel ? '' : '$'}${columnName(last.col)}`;
	if (first.col === 0 && last.col >= 0xff)
		return `${first.rowRel ? '' : '$'}${first.row + 1}:${last.rowRel ? '' : '$'}${last.row + 1}`;
	return `${cellText(first)}:${cellText(last)}`;
}

/** Decode one `row`/`col` pair; `relative` applies the cell offset to relative coordinates. */
export function readLoc(row: number, colField: number, relative: boolean, base: PtgOptions): Loc {
	const rowRel = (colField & 0x8000) !== 0;
	const colRel = (colField & 0x4000) !== 0;
	let col = colField & 0x3fff;
	let resolvedRow = row;
	if (relative) {
		if (rowRel) resolvedRow = (base.row + row) & 0xffff;
		if (colRel) col = (base.col + (colField & 0xff)) & 0xff;
	}
	return { row: resolvedRow, col, rowRel, colRel };
}

export function readArrayConstant(extra: ByteReader): string {
	const cols = extra.u8() + 1;
	const rows = extra.u16() + 1;
	const lines: string[] = [];
	for (let r = 0; r < rows; r++) {
		const values: string[] = [];
		for (let c = 0; c < cols; c++) {
			const type = extra.u8();
			if (type === 0x01) values.push(formatNumber(extra.f64()));
			else if (type === 0x02) values.push(stringLiteral(extra.xlString()));
			else if (type === 0x04) {
				values.push(extra.u8() ? 'TRUE' : 'FALSE');
				extra.skip(7);
			} else if (type === 0x10) {
				values.push(BIFF_ERRORS[extra.u8()] ?? '#N/A');
				extra.skip(7);
			} else {
				values.push('');
				extra.skip(8);
			}
		}
		lines.push(values.join(','));
	}
	return `{${lines.join(';')}}`;
}
