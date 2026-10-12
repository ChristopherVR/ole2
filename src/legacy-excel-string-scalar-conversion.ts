/** Physical LABELSST to NUMBER/BOOLERR, retaining all original SST entries. */
import type { XlsError } from './legacy-excel-workbook.js';
import { editXlsStoredCell, editXlsStoredCellWorkbookStream } from './legacy-excel-stored-cell-edit.js';
import type { XlsPreservedStringResult } from './legacy-excel-preserved-string-edit.js';

export interface XlsStringScalarEdit { row: number; col: number; value: number | boolean | XlsError; worksheetIndex?: number }
export function editXlsStringScalarWorkbookStream(input: Uint8Array, edit: XlsStringScalarEdit): XlsPreservedStringResult {
	return editXlsStoredCellWorkbookStream(input, edit, 'string');
}
export function editXlsStringScalarCell(input: Uint8Array, edit: XlsStringScalarEdit): XlsPreservedStringResult {
	try { return editXlsStoredCell(input, edit, 'string'); }
	catch { return { status: 'unchanged', bytes: input, reason: 'malformed-records' }; }
}
