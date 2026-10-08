/** Bounded absent-cell creation within an existing BIFF8 ROW/DBCELL block. */
import { model, findTarget, EditError, reject, type PhysicalRecord } from './legacy-excel-record-edit-model.js';
import { applyXlsCellRecordReplacement } from './legacy-excel-cell-record-replacement.js';
import { OLE_MAGIC } from './ole2-parser-types.js';
import { readCompoundFileStream } from './ole2-stream-edit.js';
import { resizeCompoundFileStream } from './ole2-stream-resize.js';

export interface XlsCreateNumericCell { row: number; col: number; value: number; xf: number; worksheetIndex?: number }
export type XlsCreateNumericCellFailure = 'invalid-edit' | 'unsupported-workbook' | 'malformed-records' | 'sheet-not-found' | 'cell-already-exists' | 'creation-not-supported' | 'unsupported-pointer-record' | 'container-not-writable' | 'ambiguous-cell';
export type XlsCreateNumericCellResult = {status: 'edited'; bytes: Uint8Array; recalculationRequired: true}
 | {status: 'unchanged'; bytes: Uint8Array; reason: XlsCreateNumericCellFailure};
const CELL_OPS = new Set([0x0006, 0x0201, 0x0203, 0x0204, 0x0205, 0x027e, 0x00fd, 0x00bd, 0x00be, 0x00d6]);

/** An explicit existing cell XF is required. No new rows or DIMENSIONS expansion.
 * Insertions precede no existing first-cell address: the proven numeric predecessor
 * remains at its original mapped address, preserving DBCELL first-cell semantics.
 */
export function createXlsNumericCellWorkbookStream(input: Uint8Array, supplied: XlsCreateNumericCell): XlsCreateNumericCellResult {
 const edit = {row: supplied.row, col: supplied.col, value: supplied.value, xf: supplied.xf, worksheetIndex: supplied.worksheetIndex ?? 0};
 const unchanged = (reason: XlsCreateNumericCellFailure): XlsCreateNumericCellResult => ({status: 'unchanged', bytes: input, reason});
 if (!Number.isInteger(edit.row) || edit.row < 0 || edit.row > 65535 || !Number.isInteger(edit.col) || edit.col < 0 || edit.col > 255 ||
  !Number.isFinite(edit.value) || !Number.isInteger(edit.xf) || edit.xf < 0 || edit.xf > 65535 || !Number.isInteger(edit.worksheetIndex) || edit.worksheetIndex < 0) return unchanged('invalid-edit');
 // Bound the optional creation path before allocating the full physical model.
 if (input.length > 64 * 1024 * 1024) return unchanged('creation-not-supported');
 const framing = new DataView(input.buffer, input.byteOffset, input.byteLength);
 for (let offset = 0, count = 0; offset + 4 <= input.length;) {
  if (++count > 100000) return unchanged('creation-not-supported');
  const length = framing.getUint16(offset + 2, true);
  if (length > 8224 || offset + 4 + length > input.length) return unchanged('malformed-records');
  offset += 4 + length;
 }
 try {
  const m = model(input), range = m.worksheets[edit.worksheetIndex];
  if (!range) return unchanged('sheet-not-found');
  const view = new DataView(input.buffer, input.byteOffset, input.byteLength);
  const xfs = m.records.filter(r => r.opcode === 0xe0 && r.depth === 1 && r.headerOffset < m.globalsEnd);
  const xf = xfs[edit.xf];
  if (!xf || xf.length !== 20 || (view.getUint16(xf.dataOffset + 4, true) & 4)) return unchanged('invalid-edit');
  const records = m.records.filter(r => r.depth === 1 && r.headerOffset >= range.start && r.headerOffset < range.end);
  for (const r of records.filter(r => r.opcode === 0x12)) {
   if (r.length !== 2) reject('malformed-records');
   if (view.getUint16(r.dataOffset, true) !== 0) return unchanged('creation-not-supported');
  }
  const dimensions = records.filter(r => r.opcode === 0x200);
  if (dimensions.length !== 1 || dimensions[0]!.length !== 14) reject('malformed-records');
  const dim = dimensions[0]!, firstRow = view.getUint32(dim.dataOffset, true), lastRow = view.getUint32(dim.dataOffset + 4, true), firstCol = view.getUint16(dim.dataOffset + 8, true), lastCol = view.getUint16(dim.dataOffset + 10, true);
  if (firstRow >= lastRow || lastRow > 65536 || firstCol >= lastCol || lastCol > 256) reject('malformed-records');
  if (edit.row < firstRow || edit.row >= lastRow || edit.col < firstCol || edit.col >= lastCol) return unchanged('creation-not-supported');
  for (const r of records.filter(r => [0x221, 0x4bc, 0x236, 0xe5].includes(r.opcode))) {
   if (r.opcode === 0xe5) {
    if (r.length < 2 || r.length !== 2 + view.getUint16(r.dataOffset, true) * 8) reject('malformed-records');
    for (let p = r.dataOffset + 2; p < r.dataOffset + r.length; p += 8) {
     const a = view.getUint16(p, true), b = view.getUint16(p + 2, true), c = view.getUint16(p + 4, true), d = view.getUint16(p + 6, true);
     if (a > b || c > d || d > 255) reject('malformed-records');
     if (edit.row >= a && edit.row <= b && edit.col >= c && edit.col <= d) return unchanged('creation-not-supported');
    }
   } else {
    if (r.length < 6) reject('malformed-records');
    const a = view.getUint16(r.dataOffset, true), b = view.getUint16(r.dataOffset + 2, true), c = view.getUint8(r.dataOffset + 4), d = view.getUint8(r.dataOffset + 5);
    if (a > b || c > d) reject('malformed-records');
    if (edit.row >= a && edit.row <= b && edit.col >= c && edit.col <= d) return unchanged('creation-not-supported');
   }
  }
  const cells: Array<{record: PhysicalRecord; row: number; first: number; last: number}> = [];
  for (const r of records.filter(r => CELL_OPS.has(r.opcode))) {
   if (r.length < 6) reject('malformed-records');
   const row = view.getUint16(r.dataOffset, true), first = view.getUint16(r.dataOffset + 2, true);
   let last = first;
   if (r.opcode === 0xbd || r.opcode === 0xbe) {
    const stride = r.opcode === 0xbd ? 6 : 2;
    if (r.length < 6 + stride || (r.length - 6) % stride !== 0) reject('malformed-records');
    last = view.getUint16(r.dataOffset + r.length - 2, true);
    if (last < first || last - first + 1 !== (r.length - 6) / stride) reject('malformed-records');
   }
   if (last > 255) reject('malformed-records');
   if (row === edit.row && edit.col >= first && edit.col <= last) return unchanged('cell-already-exists');
   cells.push({record: r, row, first, last});
  }
  const rowRecords = records.filter(r => r.opcode === 0x208 && r.length >= 2 && view.getUint16(r.dataOffset, true) === edit.row);
  if (rowRecords.length !== 1 || rowRecords[0]!.length !== 16) return unchanged('creation-not-supported');
  const rowRecord = rowRecords[0]!;
  const ownCells = cells.filter(c => c.row === edit.row);
  if (!ownCells.length) return unchanged('creation-not-supported');
  const rowFirst = view.getUint16(rowRecord.dataOffset + 2, true), rowLast = view.getUint16(rowRecord.dataOffset + 4, true);
  if (rowFirst >= rowLast || rowLast > 256 || ownCells.some(c => c.first < rowFirst || c.last >= rowLast)) reject('malformed-records');
  const predecessor = ownCells.filter(c => c.last < edit.col).at(-1);
  if (!predecessor || ![0x203, 0x27e, 0xbd].includes(predecessor.record.opcode)) return unchanged('creation-not-supported');
  const target = findTarget(input, m, {row: edit.row, col: predecessor.first, worksheetIndex: edit.worksheetIndex}, [0x203, 0x27e, 0xbd]);
  const blocks = records.filter(r => r.opcode === 0xd7 && r.length >= 4 && r.headerOffset - view.getUint32(r.dataOffset, true) <= rowRecord.headerOffset && rowRecord.headerOffset < r.headerOffset && target.headerOffset < r.headerOffset);
  if (blocks.length !== 1) return unchanged('creation-not-supported');
  const db = blocks[0]!, blockStart = db.headerOffset - view.getUint32(db.dataOffset, true);
  const blockRows = records.filter(r => r.opcode === 0x208 && r.headerOffset >= blockStart && r.headerOffset < db.headerOffset);
  if (!blockRows.length || blockRows.some(r => r.length !== 16)) reject('malformed-records');
  for (let i = 1; i < blockRows.length; i++) if (view.getUint16(blockRows[i]!.dataOffset, true) <= view.getUint16(blockRows[i - 1]!.dataOffset, true)) reject('malformed-records');
  const lastRowEnd = blockRows.at(-1)!.dataOffset + 16;
  const blockCells = cells.filter(c => c.record.headerOffset >= blockStart && c.record.headerOffset < db.headerOffset);
  if (blockCells.some(c => c.record.headerOffset < lastRowEnd)) return unchanged('creation-not-supported');
  for (let i = 1; i < blockCells.length; i++) {
   const a = blockCells[i - 1]!, b = blockCells[i]!;
   if (b.row < a.row || (b.row === a.row && b.first <= a.last)) reject('malformed-records');
  }
  if (target.headerOffset < lastRowEnd) return unchanged('creation-not-supported');
  const end = target.dataOffset + target.length, bytes = new Uint8Array(end - target.headerOffset + 18), number = new DataView(bytes.buffer);
  bytes.set(input.subarray(target.headerOffset, end)); const p = end - target.headerOffset;
  number.setUint16(p, 0x203, true); number.setUint16(p + 2, 14, true); number.setUint16(p + 4, edit.row, true);
  number.setUint16(p + 6, edit.col, true); number.setUint16(p + 8, edit.xf, true); number.setFloat64(p + 10, edit.value, true);
  const out = applyXlsCellRecordReplacement(input, m, {start: target.headerOffset, end, bytes});
  // ROW precedes cell storage and therefore retains its original address.
  new DataView(out.buffer).setUint16(rowRecord.dataOffset + 4, Math.max(rowLast, edit.col + 1), true);
  return {status: 'edited', bytes: out, recalculationRequired: true};
 } catch (error) {
  const reason = error instanceof EditError ? error.reason : 'malformed-records';
  return unchanged(reason === 'cell-not-supported' ? 'creation-not-supported' : reason);
 }
}

/** Bare BIFF8 or CFB; any refusal returns the original input object. */
export function createXlsNumericCell(input: Uint8Array, edit: XlsCreateNumericCell): XlsCreateNumericCellResult {
 try {
 if (!OLE_MAGIC.every((value, index) => input[index] === value)) return createXlsNumericCellWorkbookStream(input, edit);
 const streams = ['Workbook', 'Book'].map(name => ({name, bytes: readCompoundFileStream(input, [name])})).filter(s => s.bytes !== undefined);
 if (streams.length !== 1) return {status: 'unchanged', bytes: input, reason: 'unsupported-workbook'};
 const stream = streams[0]!, result = createXlsNumericCellWorkbookStream(stream.bytes!, edit);
 if (result.status !== 'edited') return {...result, bytes: input};
 const resized = resizeCompoundFileStream(input, [stream.name], result.bytes);
 return resized.ok ? {...result, bytes: resized.bytes} : {status: 'unchanged', bytes: input, reason: 'container-not-writable'};
 } catch { return {status: 'unchanged', bytes: input, reason: 'malformed-records'}; }
}
