/** MS-XLS Blank/MulBlank: replace only a physically stored formatted blank. */
import type { BiffRecord } from './legacy-excel-biff8.js';
import type { XlsCellRecordReplacement } from './legacy-excel-cell-record-splices.js';

function record(opcode: number, data: Uint8Array): Uint8Array {
 const out = new Uint8Array(4 + data.length), view = new DataView(out.buffer);
 view.setUint16(0, opcode, true); view.setUint16(2, data.length, true); out.set(data, 4);
 return out;
}

/** Stage a NUMBER, retaining the selected ixfe and all other packed ixfe bytes. */
export function replaceXlsBlankRecord(input: Uint8Array, target: BiffRecord, col: number, value: number): XlsCellRecordReplacement {
 if (!Number.isInteger(col) || col < 0 || col > 255 || !Number.isFinite(value) ||
  !Number.isSafeInteger(target.headerOffset) || target.headerOffset < 0 || target.dataOffset !== target.headerOffset + 4 ||
  !Number.isInteger(target.length) || target.length < 6 || target.dataOffset + target.length > input.length ||
  ![0x0201, 0x00be].includes(target.opcode)) throw new Error('Invalid blank record replacement');
 const view = new DataView(input.buffer, input.byteOffset, input.byteLength), p = target.dataOffset;
 if (view.getUint16(target.headerOffset, true) !== target.opcode || view.getUint16(target.headerOffset + 2, true) !== target.length)
  throw new Error('Blank record framing mismatch');
 const packed = target.opcode === 0x00be;
 if (packed ? target.length < 10 || (target.length - 6) % 2 !== 0 : target.length !== 6)
  throw new Error('Invalid blank record length');
 const row = view.getUint16(p, true), first = view.getUint16(p + 2, true), count = packed ? (target.length - 6) / 2 : 1;
 const selected = col - first;
 if (first > 255 || first + count - 1 > 255 || selected < 0 || selected >= count ||
  (packed && view.getUint16(p + target.length - 2, true) !== first + count - 1)) throw new Error('Invalid blank column range');
 const data = new Uint8Array(14), cell = new DataView(data.buffer);
 cell.setUint16(0, row, true); cell.setUint16(2, col, true);
 cell.setUint16(4, view.getUint16(p + 4 + (packed ? selected * 2 : 0), true), true);
 cell.setFloat64(6, value, true);
 const number = record(0x0203, data);
 const group = (start: number, size: number): Uint8Array => {
  if (!size) return new Uint8Array();
  const data = new Uint8Array(size === 1 ? 6 : 6 + size * 2), groupView = new DataView(data.buffer);
  groupView.setUint16(0, row, true); groupView.setUint16(2, first + start, true);
  data.set(input.subarray(p + 4 + start * 2, p + 4 + (start + size) * 2), 4);
  if (size > 1) groupView.setUint16(data.length - 2, first + start + size - 1, true);
  return record(size === 1 ? 0x0201 : 0x00be, data);
 };
 const prefix = packed ? group(0, selected) : new Uint8Array(), suffix = packed ? group(selected + 1, count - selected - 1) : new Uint8Array();
 const bytes = new Uint8Array(prefix.length + number.length + suffix.length);
 bytes.set(prefix); bytes.set(number, prefix.length); bytes.set(suffix, prefix.length + number.length);
 return { start: target.headerOffset, end: p + target.length, bytes };
}
