/** Checked stored-cell conversions over BLANK/MULBLANK or BOOLERR records.
 * The parsed model and record replacement never form a public unchecked API.
 */
import type { XlsCellValue } from './legacy-excel-workbook.js';
import { model, findTarget, reject, EditError } from './legacy-excel-record-edit-model.js';
import { replaceXlsBlankRecord } from './legacy-excel-blank-record.js';
import { type XlsCellRecordReplacement } from './legacy-excel-cell-record-splices.js';
import { applyXlsCellRecordReplacement } from './legacy-excel-cell-record-replacement.js';
import { editXlsStringWorkbookStream, type XlsPreservedStringResult } from './legacy-excel-preserved-string-edit.js';
import { editXlsBoolErrorWorkbookStream, normalizeXlsErrorValue, isValidXlsBoolErrorRecordValue } from './legacy-excel-bool-error-edit.js';
import { OLE_MAGIC } from './ole2-parser-types.js';
import { readCompoundFileStream, replaceCompoundFileStream } from './ole2-stream-edit.js';
import { resizeCompoundFileStream } from './ole2-stream-resize.js';

export type XlsStoredCellEdit = { row: number; col: number; value: Exclude<XlsCellValue, null>; worksheetIndex?: number };
type Source = 'blank' | 'bool-error';

/** Every source and coordinate is validated against original physical records. */
export function editXlsStoredCellWorkbookStream(input: Uint8Array, edit: XlsStoredCellEdit, sourceKind: Source): XlsPreservedStringResult {
 const unchanged = (reason: Extract<XlsPreservedStringResult, { status: 'unchanged' }>['reason']): XlsPreservedStringResult => ({ status: 'unchanged', bytes: input, reason });
 // Capture caller values once; subsequent checked editors receive ordinary data.
 const row = edit.row, col = edit.col, worksheetIndex = edit.worksheetIndex ?? 0, supplied = edit.value;
 const value = supplied !== null && typeof supplied === 'object' ? normalizeXlsErrorValue(supplied) : supplied;
 if (!['blank','bool-error'].includes(sourceKind) || !Number.isInteger(row) || row < 0 || row > 65535 || !Number.isInteger(col) || col < 0 || col > 255 || !Number.isInteger(worksheetIndex) || worksheetIndex < 0 ||
  (sourceKind === 'bool-error' && typeof value !== 'number' && typeof value !== 'string') ||
  (typeof value === 'number' ? !Number.isFinite(value) : typeof value !== 'string' && typeof value !== 'boolean' && (value === undefined || value === null || typeof value !== 'object'))) return unchanged('invalid-edit');
 try {
  const m = model(input), target = findTarget(input, m, { row, col, worksheetIndex }, sourceKind === 'blank' ? [0x0201, 0x00be] : [0x0205]);
  const range = m.worksheets[worksheetIndex]!, source = new DataView(input.buffer, input.byteOffset, input.byteLength);
  for (const record of m.records.filter(r => r.opcode === 0x00e5 && r.depth === 1 && r.headerOffset >= range.start && r.headerOffset < range.end)) {
   if (record.length < 2) reject('malformed-records');
   const count = source.getUint16(record.dataOffset, true);
   if (record.length !== 2 + count * 8) reject('malformed-records');
   for (let p = record.dataOffset + 2; p < record.dataOffset + record.length; p += 8) {
    const firstRow = source.getUint16(p, true), lastRow = source.getUint16(p + 2, true), firstCol = source.getUint16(p + 4, true), lastCol = source.getUint16(p + 6, true);
    if (firstRow > lastRow || firstCol > lastCol || lastCol > 255) reject('malformed-records');
    if (row >= firstRow && row <= lastRow && col >= firstCol && col <= lastCol && (row !== firstRow || col !== firstCol)) reject('cell-not-supported');
   }
  }
  // A valid temporary NUMBER lets the existing checked SST/BOOLERR writers reuse
  // their pointer and alias guards. Intermediate bytes never leave this call.
  let replacement: XlsCellRecordReplacement;
  if (sourceKind === 'blank') replacement = replaceXlsBlankRecord(input, target, col, typeof value === 'number' ? value : 0);
  else {
   if (!isValidXlsBoolErrorRecordValue(source.getUint8(target.dataOffset + 7), source.getUint8(target.dataOffset + 6))) reject('malformed-records');
   const bytes = new Uint8Array(18), number = new DataView(bytes.buffer);
   number.setUint16(0,0x0203,true); number.setUint16(2,14,true);
   bytes.set(input.subarray(target.dataOffset,target.dataOffset+6),4);
   number.setFloat64(10,typeof value === 'number' ? value : 0,true);
   replacement = { start: target.headerOffset, end: target.dataOffset + target.length, bytes };
  }
  const staged = applyXlsCellRecordReplacement(input, m, replacement);
  if (typeof value === 'number') return { status: 'edited', bytes: staged, recalculationRequired: true };
  const result = typeof value === 'string' ? editXlsStringWorkbookStream(staged, { row, col, worksheetIndex, value })
   : editXlsBoolErrorWorkbookStream(staged, { row, col, worksheetIndex, value: value! });
  return result.status === 'edited' ? result : { ...result, bytes: input };
 } catch (error) { return unchanged(error instanceof EditError ? error.reason : 'malformed-records'); }
}

/** Bare BIFF8 or CFB. Failures return the exact original input object. */
export function editXlsStoredCell(input: Uint8Array, edit: XlsStoredCellEdit, sourceKind: Source): XlsPreservedStringResult {
 if (!OLE_MAGIC.every((value, index) => input[index] === value)) return editXlsStoredCellWorkbookStream(input, edit, sourceKind);
 const streams = ['Workbook', 'Book'].map(name => ({ name, bytes: readCompoundFileStream(input, [name]) })).filter(s => s.bytes !== undefined);
 if (streams.length !== 1) return { status: 'unchanged', bytes: input, reason: 'unsupported-workbook' };
 const stream = streams[0]!, result = editXlsStoredCellWorkbookStream(stream.bytes!, edit, sourceKind);
 if (result.status !== 'edited') return { ...result, bytes: input };
 if (result.bytes.length === stream.bytes!.length) {
  const bytes = replaceCompoundFileStream(input, [stream.name], result.bytes);
  return bytes === input ? { status: 'unchanged', bytes: input, reason: 'container-not-writable' } : { ...result, bytes };
 }
 const resized = resizeCompoundFileStream(input, [stream.name], result.bytes);
 return resized.ok ? { ...result, bytes: resized.bytes } : { status: 'unchanged', bytes: input, reason: 'container-not-writable' };
}
