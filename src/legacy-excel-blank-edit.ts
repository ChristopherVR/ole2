/** Existing BLANK/MULBLANK value writes. No implicit cells or rows are created. */
import { editXlsStoredCell, editXlsStoredCellWorkbookStream, type XlsStoredCellEdit } from './legacy-excel-stored-cell-edit.js';
import type { XlsPreservedStringResult } from './legacy-excel-preserved-string-edit.js';
export function editXlsBlankWorkbookStream(input: Uint8Array, edit: XlsStoredCellEdit): XlsPreservedStringResult {
 return editXlsStoredCellWorkbookStream(input, edit, 'blank');
}
export function editXlsBlankCell(input: Uint8Array, edit: XlsStoredCellEdit): XlsPreservedStringResult {
 return editXlsStoredCell(input, edit, 'blank');
}
