/** Lossless RK/MULRK promotion to NUMBER using checked record relocation. */
import type { BiffRecord } from './legacy-excel-biff8.js';
import { model, findTarget, reject } from './legacy-excel-record-edit-model.js';
import { applyXlsCellRecordReplacement } from './legacy-excel-cell-record-replacement.js';

function promote(input: Uint8Array, target: BiffRecord, col: number, value: number): Uint8Array {
 const source = new DataView(input.buffer, input.byteOffset, input.byteLength), p = target.dataOffset;
 const packed = target.opcode === 0x00bd;
 const first = packed ? source.getUint16(p + 2, true) : col;
 const count = packed ? (target.length - 6) / 6 : 1, selected = col - first;
 const record = (opcode: number, data: Uint8Array): Uint8Array => {
  const out = new Uint8Array(data.length + 4), view = new DataView(out.buffer);
  view.setUint16(0, opcode, true); view.setUint16(2, data.length, true); out.set(data, 4); return out;
 };
 const group = (start: number, size: number): Uint8Array => {
  if (!size) return new Uint8Array();
  const data = new Uint8Array(size === 1 ? 10 : 6 + size * 6), view = new DataView(data.buffer);
  view.setUint16(0, source.getUint16(p, true), true); view.setUint16(2, first + start, true);
  data.set(input.subarray(p + 4 + start * 6, p + 4 + (start + size) * 6), 4);
  if (size > 1) view.setUint16(data.length - 2, first + start + size - 1, true);
  return record(size === 1 ? 0x027e : 0x00bd, data);
 };
 const data = new Uint8Array(14), view = new DataView(data.buffer);
 view.setUint16(0, source.getUint16(p, true), true); view.setUint16(2, col, true);
 view.setUint16(4, source.getUint16(p + (packed ? 4 + selected * 6 : 4), true), true);
 view.setFloat64(6, value, true);
 const prefix = packed ? group(0, selected) : new Uint8Array(), number = record(0x0203, data);
 const suffix = packed ? group(selected + 1, count - selected - 1) : new Uint8Array();
 const bytes = new Uint8Array(prefix.length + number.length + suffix.length);
 bytes.set(prefix); bytes.set(number, prefix.length); bytes.set(suffix, prefix.length + number.length);
 return bytes;
}

/** Only complete BIFF8 workbooks with a checked tab directory can grow. */
export function promoteXlsNumericCell(input: Uint8Array, edit: {row: number; col: number; value: number; worksheetIndex?: number}): Uint8Array {
 const m = model(input), target = findTarget(input, m, edit, [0x027e, 0x00bd]);
 const source = new DataView(input.buffer, input.byteOffset, input.byteLength), range = m.worksheets[edit.worksheetIndex ?? 0]!;
 // A non-anchor merged cell cannot be made into an independent value.
 for (const record of m.records.filter(r => r.opcode === 0x00e5 && r.depth === 1 && r.headerOffset >= range.start && r.headerOffset < range.end)) {
  if (record.length < 2 || record.length !== 2 + source.getUint16(record.dataOffset, true) * 8) reject('malformed-records');
  for (let p = record.dataOffset + 2; p < record.dataOffset + record.length; p += 8) {
   const firstRow = source.getUint16(p, true), lastRow = source.getUint16(p + 2, true), firstCol = source.getUint16(p + 4, true), lastCol = source.getUint16(p + 6, true);
   if (firstRow > lastRow || firstCol > lastCol || lastCol > 255) reject('malformed-records');
   if (edit.row >= firstRow && edit.row <= lastRow && edit.col >= firstCol && edit.col <= lastCol && (edit.row !== firstRow || edit.col !== firstCol)) reject('cell-not-supported');
  }
 }
 return applyXlsCellRecordReplacement(input, m, {start: target.headerOffset, end: target.dataOffset + target.length, bytes: promote(input, target, edit.col, edit.value)});
}
