/** Existing BLANK/MULBLANK value writes. No implicit cells or rows are created. */
import type { XlsCellValue } from './legacy-excel-workbook.js';
import { model, findTarget, reject, EditError, type Model } from './legacy-excel-record-edit-model.js';
import { replaceXlsBlankRecord } from './legacy-excel-blank-record.js';
import { relocateXlsDbCells, type XlsCellRecordReplacement } from './legacy-excel-cell-record-splices.js';
import { KNOWN_XLS_RECORDS, XLS_FRT_WRAPPER } from './legacy-excel-known-records.js';
import { editXlsStringWorkbookStream, type XlsPreservedStringResult } from './legacy-excel-preserved-string-edit.js';
import { editXlsBoolErrorWorkbookStream, normalizeXlsErrorValue } from './legacy-excel-bool-error-edit.js';
import { OLE_MAGIC } from './ole2-parser-types.js';
import { readCompoundFileStream, replaceCompoundFileStream } from './ole2-stream-edit.js';
import { resizeCompoundFileStream } from './ole2-stream-resize.js';

type Edit = { row: number; col: number; value: Exclude<XlsCellValue, null>; worksheetIndex?: number };

/** Private relocation over this call's checked model and generated cell bytes. */
function applyXlsCellRecordReplacement(input: Uint8Array, m: Model, replacement: XlsCellRecordReplacement): Uint8Array {
 if (!Number.isSafeInteger(replacement.start) || !Number.isSafeInteger(replacement.end) || replacement.start < 0 || replacement.end <= replacement.start || replacement.end > input.length || !(replacement.bytes instanceof Uint8Array)) reject('malformed-records');
 const target = m.records.find(r => r.headerOffset === replacement.start);
 if (!target || target.dataOffset + target.length !== replacement.end || ![0x0201,0x00be].includes(target.opcode)) reject('malformed-records');
 const delta = replacement.bytes.length - (replacement.end - replacement.start);
 if (delta !== 0 && m.records.some(r => !KNOWN_XLS_RECORDS.has(r.opcode) || r.opcode === XLS_FRT_WRAPPER)) reject('unsupported-pointer-record');
 const map = (address: number): number => {
  if (!Number.isSafeInteger(address) || address < 0 || address > input.length || (address > replacement.start && address < replacement.end)) reject('malformed-records');
  return address >= replacement.end ? address + delta : address;
 };
 const out = new Uint8Array(input.length + delta);
 out.set(input.subarray(0, replacement.start)); out.set(replacement.bytes, replacement.start);
 out.set(input.subarray(replacement.end), replacement.start + replacement.bytes.length);
 if (delta !== 0) {
  relocateXlsDbCells(input, out, m.records, replacement, map);
  const source = new DataView(input.buffer, input.byteOffset, input.byteLength), dest = new DataView(out.buffer), byOffset = new Map(m.records.map(r => [r.headerOffset, r]));
  for (const r of m.records) {
   if (r.opcode === 0x0085) {
    if (r.headerOffset >= m.globalsEnd || r.depth !== 1) reject('malformed-records');
    dest.setUint32(map(r.dataOffset), map(source.getUint32(r.dataOffset, true)), true);
   }
   if (r.opcode === 0x020b) {
    if (r.length < 16 || (r.length - 16) % 4 !== 0) reject('malformed-records');
    const owner = m.worksheets.find(sheet => r.depth === 1 && r.headerOffset >= sheet.start && r.headerOffset < sheet.end);
    if (!owner) reject('malformed-records');
    const pointers = new Set<number>();
    for (let p = r.dataOffset + 12; p < r.dataOffset + r.length; p += 4) {
     const address = source.getUint32(p, true);
     const pointed = byOffset.get(address);
     if (pointed?.opcode !== (p === r.dataOffset + 12 ? 0x0055 : 0x00d7) || pointed.depth !== 1 || address < owner.start || address >= owner.end) reject('malformed-records');
     if (pointers.has(address)) reject('malformed-records');
     pointers.add(address);
     dest.setUint32(map(p), map(address), true);
    }
   }
  }
 }
 return out;
}
export function editXlsBlankWorkbookStream(input: Uint8Array, edit: Edit): XlsPreservedStringResult {
 const unchanged = (reason: Extract<XlsPreservedStringResult, { status: 'unchanged' }>['reason']): XlsPreservedStringResult => ({ status: 'unchanged', bytes: input, reason });
 // Capture caller values once; subsequent checked editors receive ordinary data.
 const row = edit.row, col = edit.col, worksheetIndex = edit.worksheetIndex ?? 0, supplied = edit.value;
 const value = supplied !== null && typeof supplied === 'object' ? normalizeXlsErrorValue(supplied) : supplied;
 if (!Number.isInteger(row) || row < 0 || row > 65535 || !Number.isInteger(col) || col < 0 || col > 255 || !Number.isInteger(worksheetIndex) || worksheetIndex < 0 ||
  (typeof value === 'number' ? !Number.isFinite(value) : typeof value !== 'string' && typeof value !== 'boolean' && (value === undefined || value === null || typeof value !== 'object'))) return unchanged('invalid-edit');
 try {
  const m = model(input), target = findTarget(input, m, { row, col, worksheetIndex }, [0x0201, 0x00be]);
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
  const staged = applyXlsCellRecordReplacement(input, m, replaceXlsBlankRecord(input, target, col, typeof value === 'number' ? value : 0));
  if (typeof value === 'number') return { status: 'edited', bytes: staged, recalculationRequired: true };
  const result = typeof value === 'string' ? editXlsStringWorkbookStream(staged, { row, col, worksheetIndex, value })
   : editXlsBoolErrorWorkbookStream(staged, { row, col, worksheetIndex, value: value! });
  return result.status === 'edited' ? result : { ...result, bytes: input };
 } catch (error) { return unchanged(error instanceof EditError ? error.reason : 'malformed-records'); }
}

/** Bare BIFF8 or CFB. Failures return the exact original input object. */
export function editXlsBlankCell(input: Uint8Array, edit: Edit): XlsPreservedStringResult {
 if (!OLE_MAGIC.every((value, index) => input[index] === value)) return editXlsBlankWorkbookStream(input, edit);
 const streams = ['Workbook', 'Book'].map(name => ({ name, bytes: readCompoundFileStream(input, [name]) })).filter(s => s.bytes !== undefined);
 if (streams.length !== 1) return { status: 'unchanged', bytes: input, reason: 'unsupported-workbook' };
 const stream = streams[0]!, result = editXlsBlankWorkbookStream(stream.bytes!, edit);
 if (result.status !== 'edited') return { ...result, bytes: input };
 if (result.bytes.length === stream.bytes!.length) {
  const bytes = replaceCompoundFileStream(input, [stream.name], result.bytes);
  return bytes === input ? { status: 'unchanged', bytes: input, reason: 'container-not-writable' } : { ...result, bytes };
 }
 const resized = resizeCompoundFileStream(input, [stream.name], result.bytes);
 return resized.ok ? { ...result, bytes: resized.bytes } : { status: 'unchanged', bytes: input, reason: 'container-not-writable' };
}
