/** Existing physical BOOLERR to NUMBER/plain LABELSST, preserving its ixfe. */
import { editXlsStoredCell, editXlsStoredCellWorkbookStream } from './legacy-excel-stored-cell-edit.js';
import type { XlsPreservedStringResult } from './legacy-excel-preserved-string-edit.js';
type Edit = { row: number; col: number; value: number | string; worksheetIndex?: number };
export function editXlsBoolErrorScalarWorkbookStream(input: Uint8Array, edit: Edit): XlsPreservedStringResult {
 return editXlsStoredCellWorkbookStream(input, edit, 'bool-error');
}
export function editXlsBoolErrorScalarCell(input: Uint8Array, edit: Edit): XlsPreservedStringResult {
 return editXlsStoredCell(input, edit, 'bool-error');
}
