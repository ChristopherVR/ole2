/** MS-XLS 2.4.24 BoolErr and 2.5.10 Bes: checked non-formula value edits.
 * NUMBER/RK/MULRK become BOOLERR without rebuilding worksheet rows. String
 * cells are refused: this editor does not remove SST references.
 */
import type { XlsError, XlsErrorCode } from './legacy-excel-workbook.js';
import { model, findTarget, reject, EditError } from './legacy-excel-record-edit-model.js';
import { replaceXlsNumericRecord, relocateXlsDbCells } from './legacy-excel-cell-record-splices.js';
import { readRecords } from './legacy-excel-biff8.js';
import { KNOWN_XLS_RECORDS, XLS_FRT_WRAPPER } from './legacy-excel-known-records.js';
import { OLE_MAGIC } from './ole2-parser-types.js';
import { readCompoundFileStream, replaceCompoundFileStream } from './ole2-stream-edit.js';
import { resizeCompoundFileStream } from './ole2-stream-resize.js';
import type { XlsPreservedStringResult } from './legacy-excel-preserved-string-edit.js';

const ERRORS: Readonly<Record<XlsErrorCode, number>> = Object.freeze({
 '#NULL!': 0, '#DIV/0!': 7, '#VALUE!': 15, '#REF!': 23,
 '#NAME?': 29, '#NUM!': 36, '#N/A': 42, '#GETTING_DATA': 43,
});
export type XlsBoolErrorValue = boolean | XlsError;
/** MS-XLS Bes: discriminate only defined boolean bytes or native error codes. */
export function isValidXlsBoolErrorRecordValue(flag: number, value: number): boolean {
 return (flag === 0 && (value === 0 || value === 1)) || (flag === 1 && Object.values(ERRORS).includes(value));
}
/** Capture accessor-backed inputs once; later serialization uses plain data. */
export function normalizeXlsErrorValue(value: unknown): XlsError | undefined {
 try {
  if (!value || typeof value !== 'object' || Array.isArray(value)) return undefined;
  const keys = Object.keys(value);
  if (keys.length !== 1 || keys[0] !== 'error') return undefined;
  const error = (value as XlsError).error;
  return typeof error === 'string' && Object.prototype.hasOwnProperty.call(ERRORS, error) ? { error } : undefined;
 } catch { return undefined; }
}
export function isXlsErrorValue(value: unknown): value is XlsError {
 return normalizeXlsErrorValue(value) !== undefined;
}
type Edit = { row: number; col: number; value: XlsBoolErrorValue; worksheetIndex?: number };

/** Checked bare Workbook stream editor; result failures retain the input bytes. */
export function editXlsBoolErrorWorkbookStream(input: Uint8Array, edit: Edit): XlsPreservedStringResult {
 const unchanged = (reason: Extract<XlsPreservedStringResult, { status: 'unchanged' }>['reason']): XlsPreservedStringResult => ({ status: 'unchanged', bytes: input, reason });
 const suppliedValue = edit.value;
 const value = typeof suppliedValue === 'boolean' ? suppliedValue : normalizeXlsErrorValue(suppliedValue);
 if (!Number.isInteger(edit.row) || edit.row < 0 || edit.row > 65535 || !Number.isInteger(edit.col) || edit.col < 0 || edit.col > 255 ||
  !Number.isInteger(edit.worksheetIndex ?? 0) || (edit.worksheetIndex ?? 0) < 0 ||
  value === undefined) return unchanged('invalid-edit');
 try {
  const m = model(input), target = findTarget(input, m, edit, [0x0203, 0x027e, 0x00bd, 0x0205]);
  const source = new DataView(input.buffer, input.byteOffset, input.byteLength);
  if (target.opcode === 0x0205) {
   const flag = source.getUint8(target.dataOffset + 7), value = source.getUint8(target.dataOffset + 6);
   if (!isValidXlsBoolErrorRecordValue(flag, value)) reject('malformed-records');
  }
  const data = new Uint8Array(12), view = new DataView(data.buffer);
  view.setUint16(0, 0x0205, true); view.setUint16(2, 8, true);
  view.setUint16(4, edit.row, true); view.setUint16(6, edit.col, true);
  const xfOffset = target.dataOffset + (target.opcode === 0x00bd ? 4 + (edit.col - source.getUint16(target.dataOffset + 2, true)) * 6 : 4);
  view.setUint16(8, source.getUint16(xfOffset, true), true);
  data[10] = typeof value === 'boolean' ? Number(value) : ERRORS[value.error];
  data[11] = typeof value === 'boolean' ? 0 : 1;
  let replacement = { start: target.headerOffset, end: target.dataOffset + target.length, bytes: data };
  if (target.opcode === 0x0203 || target.opcode === 0x00bd) {
   const numeric = replaceXlsNumericRecord(input, target, edit.col, 0);
   const label = readRecords(numeric.bytes, 0, numeric.bytes.length).find(r => r.opcode === 0x00fd)!;
   const bytes = new Uint8Array(numeric.bytes.length - 2);
   bytes.set(numeric.bytes.subarray(0, label.headerOffset));
   bytes.set(data, label.headerOffset);
   bytes.set(numeric.bytes.subarray(label.dataOffset + label.length), label.headerOffset + data.length);
   replacement = { ...numeric, bytes };
  }
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
   const dest = new DataView(out.buffer), byOffset = new Map(m.records.map(r => [r.headerOffset, r]));
   for (const r of m.records) {
    if (r.opcode === 0x0085) {
     if (r.headerOffset >= m.globalsEnd || r.depth !== 1) reject('malformed-records');
     dest.setUint32(map(r.dataOffset), map(source.getUint32(r.dataOffset, true)), true);
    }
    if (r.opcode === 0x020b) {
     if (r.length < 16 || (r.length - 16) % 4 !== 0) reject('malformed-records');
     for (let p = r.dataOffset + 12; p < r.dataOffset + r.length; p += 4) {
      const address = source.getUint32(p, true);
      if (byOffset.get(address)?.opcode !== (p === r.dataOffset + 12 ? 0x0055 : 0x00d7)) reject('malformed-records');
      dest.setUint32(map(p), map(address), true);
     }
    }
   }
  }
  return { status: 'edited', bytes: out, recalculationRequired: true };
 } catch (error) { return unchanged(error instanceof EditError ? error.reason : 'malformed-records'); }
}

/** Whole-CFB or bare BIFF8 edit. Every unrelated stream/allocation is preserved. */
export function editXlsBoolErrorCell(input: Uint8Array, edit: Edit): XlsPreservedStringResult {
 if (!OLE_MAGIC.every((value, index) => input[index] === value)) return editXlsBoolErrorWorkbookStream(input, edit);
 const streams = ['Workbook', 'Book'].map(name => ({ name, bytes: readCompoundFileStream(input, [name]) })).filter(s => s.bytes !== undefined);
 if (streams.length !== 1) return { status: 'unchanged', bytes: input, reason: 'unsupported-workbook' };
 const stream = streams[0]!, result = editXlsBoolErrorWorkbookStream(stream.bytes!, edit);
 if (result.status !== 'edited') return { ...result, bytes: input };
 if (result.bytes.length === stream.bytes!.length) {
  const bytes = replaceCompoundFileStream(input, [stream.name], result.bytes);
  return bytes === input ? { status: 'unchanged', bytes: input, reason: 'container-not-writable' } : { ...result, bytes };
 }
 const resized = resizeCompoundFileStream(input, [stream.name], result.bytes);
 return resized.ok ? { ...result, bytes: resized.bytes } : { status: 'unchanged', bytes: input, reason: 'container-not-writable' };
}
