/**
 * Formula cells of a BIFF8 worksheet: the `FORMULA` cached result, the
 * `SHRFMLA` / `ARRAY` definitions a `PtgExp` token points at, and decoding of
 * every collected formula once the whole substream has been read (shared
 * definitions follow the first cell that uses them).
 *
 * @module legacy-excel-workbook-formulas
 */
import type { XlsGlobals } from './legacy-excel-workbook-globals.js';
import { BIFF_ERRORS, decodeFormula } from './legacy-excel-workbook-ptg.js';
import { ByteReader } from './legacy-excel-workbook-records.js';
import type { XlsCell, XlsRange } from './legacy-excel-workbook-types.js';

export interface PendingFormula {
	cell: XlsCell;
	rgce: Uint8Array;
	extra: Uint8Array;
}

export interface FormulaDefinition {
	range: XlsRange;
	rgce: Uint8Array;
	extra: Uint8Array;
}

function readRefU(r: ByteReader): XlsRange {
	return { firstRow: r.u16(), lastRow: r.u16(), firstCol: r.u8(), lastCol: r.u8() };
}

export function formulaDefinition(data: Uint8Array, headerBytes: number): FormulaDefinition {
	const r = new ByteReader(data);
	const range = readRefU(r);
	r.skip(headerBytes);
	const cce = r.u16();
	return {
		range,
		rgce: data.subarray(r.pos, r.pos + cce),
		extra: data.subarray(r.pos + cce),
	};
}

export function formulaValue(data: Uint8Array): XlsCell['value'] | 'string' {
	if (data[12] === 0xff && data[13] === 0xff) {
		const type = data[6];
		if (type === 0) return 'string';
		if (type === 1) return data[8] !== 0;
		if (type === 2) return { error: BIFF_ERRORS[data[8] ?? 0] ?? '#N/A' };
		return '';
	}
	return new ByteReader(data, 6).f64();
}

export function resolveFormulas(
	formulas: PendingFormula[],
	shared: Map<string, FormulaDefinition>,
	arrays: Map<string, FormulaDefinition>,
	globals: XlsGlobals,
): void {
	for (const { cell, rgce, extra } of formulas) {
		let text: string | undefined;
		if (rgce.length === 5 && rgce[0] === 0x01) {
			const anchor = new ByteReader(rgce, 1);
			const key = `${anchor.u16()},${anchor.u16()}`;
			const array = arrays.get(key);
			const definition = shared.get(key);
			if (array) {
				if (array.range.firstRow !== cell.row || array.range.firstCol !== cell.col) continue;
				cell.arrayRange = array.range;
				text = decodeFormula(array.rgce, array.extra, globals.context, cell);
			} else if (definition) {
				cell.sharedFormula = true;
				text = decodeFormula(definition.rgce, definition.extra, globals.context, {
					row: cell.row,
					col: cell.col,
					shared: true,
				});
			}
		} else text = decodeFormula(rgce, extra, globals.context, cell);
		if (text === undefined) cell.formulaUndecoded = true;
		else cell.formula = text;
	}
}
