import { readFileSync } from 'node:fs';
import { describe, expect, it } from 'vitest';
import { editXlsNumericCell, encodeRkExact } from '../src/legacy-excel-biff8-edit.js';
import { decodeRk, findAllWorksheetRanges, readRecords, writeOleXlsNumericCellEdit } from '../src/legacy-excel-biff8.js';
import { editXlsStringCell, writeOleXlsStringCellEdit } from '../src/legacy-excel-biff8-writer.js';
import { unwrapXlsBytes } from '../src/legacy-excel-cfb.js';
import { readXlsWorkbook } from '../src/legacy-excel-workbook.js';

const fixture = (name = 'workbook-features.xls') => new Uint8Array(readFileSync(new URL(`./fixtures/xls/${name}`, import.meta.url)));
const cell = (bytes: Uint8Array, sheet: number, row: number, col: number) => readXlsWorkbook(bytes).sheets[sheet]?.cells.find(c => c.row === row && c.col === col);

describe('exact BIFF8 numeric editing on Excel-generated fixtures', () => {
 it('edits a MULRK cell while preserving every byte outside its four-byte value', () => {
  const input = fixture();
  const stream = unwrapXlsBytes(input).workbookBytes;
  const range = findAllWorksheetRanges(stream)[0]!;
  const v = new DataView(stream.buffer, stream.byteOffset, stream.byteLength);
  const record = readRecords(stream, range.start, range.end).find(r => r.opcode === 0xbd && v.getUint16(r.dataOffset, true) === 1 && v.getUint16(r.dataOffset + 2, true) === 1)!;
  expect(record).toBeDefined();
  const result = editXlsNumericCell(input, { row: 1, col: 1, value: 2.75 });
  expect(result.status).toBe('edited');
  expect(result.bytes.length).toBe(input.length);
  expect(cell(result.bytes, 0, 1, 1)?.value).toBe(2.75);
  expect(cell(result.bytes, 0, 1, 2)?.value).toBe(10);
  const actual = unwrapXlsBytes(result.bytes).workbookBytes;
  const masked = actual.slice();
  masked.set(stream.subarray(record.dataOffset + 6, record.dataOffset + 10), record.dataOffset + 6);
  expect(masked).toEqual(stream);
  expect(readXlsWorkbook(result.bytes).sheets.slice(1)).toEqual(readXlsWorkbook(input).sheets.slice(1));
  expect(result).toMatchObject({ recalculationRequired: true });
 });
 it('edits an existing cell on a later hidden worksheet', () => {
  const input = fixture();
  const before = readXlsWorkbook(input);
  const target = before.sheets[2]!.cells.find(c => typeof c.value === 'number' && !c.formula)!;
  expect(target).toBeDefined();
  const result = editXlsNumericCell(input, { worksheetIndex: 2, row: target.row, col: target.col, value: 43 });
  expect(result.status).toBe('edited');
  expect(cell(result.bytes, 2, target.row, target.col)?.value).toBe(43);
  expect(readXlsWorkbook(result.bytes).sheets[0]).toEqual(before.sheets[0]);
 });
 it('promotes an unrepresentable RK value without rounding and does not replace a formula', () => {
  const input = fixture();
  const promoted = editXlsNumericCell(input, { row: 1, col: 1, value: Math.PI });
  expect(promoted.status).toBe('edited');
  expect(cell(promoted.bytes, 0, 1, 1)?.value).toBe(Math.PI);
  expect(cell(writeOleXlsNumericCellEdit(input, { row: 1, col: 1, value: Math.PI }), 0, 1, 1)?.value).toBe(Math.PI);
  expect(editXlsNumericCell(input, { worksheetIndex: 1, row: 0, col: 0, value: 9 })).toMatchObject({ status: 'unchanged', reason: 'cell-not-supported' });
 });
 it('rejects invalid coordinates, protected files and truncated records', () => {
  const input = fixture();
  for (const edit of [{row: -1,col:0,value:1}, {row:0,col:256,value:1}, {row:0.1,col:0,value:1}, {row:0,col:0,value:Infinity}]) expect(editXlsNumericCell(input, edit)).toMatchObject({status:'unchanged',reason:'invalid-edit'});
  expect(editXlsNumericCell(fixture('workbook-encrypted.xls'), {row:0,col:0,value:1})).toMatchObject({status:'unchanged',reason:'unsupported-workbook'});
  const stream = unwrapXlsBytes(input).workbookBytes;
  expect(editXlsNumericCell(stream.subarray(0, findAllWorksheetRanges(stream)[0]!.start + 7), {row:1,col:1,value:1})).toMatchObject({status:'unchanged',reason:'malformed-records'});
 });
 it('refuses string resizing that would discard the real workbook cell table', () => {
  const input = fixture('workbook-styles.xls');
  expect(writeOleXlsStringCellEdit(input, {row:0,col:0,value:'A fresh label'})).toBe(input);
  expect(editXlsStringCell(input, {row:0,col:0,value:'A fresh label'})).toMatchObject({status:'unchanged',reason:'unsupported-or-unsafe-edit',bytes:input});
 });
});

describe('RK exact encoding', () => {
 it('retains integer boundaries, scaled decimals, binary fractions and signed zero', () => {
  for (const n of [-536870912, 536870911, -3.14, 2.75, 0, -0, 1.5, 536870912]) {
   const encoded = encodeRkExact(n);
   expect(encoded).toBeDefined();
   expect(Object.is(decodeRk(encoded!), n)).toBe(true);
  }
  expect(encodeRkExact(Math.PI)).toBeUndefined();
  expect(encodeRkExact(Infinity)).toBeUndefined();
 });
});

function rec(opcode: number, data: number[]): number[] {
 return [opcode & 255, opcode >> 8, data.length & 255, data.length >> 8, ...data];
}
function u32(n: number): number[] { return [n & 255, (n >>> 8) & 255, (n >>> 16) & 255, n >>> 24]; }
function testBook(duplicate?: number): Uint8Array {
 const bof = (kind: number) => rec(0x809,[0,6,kind,0,...new Array(12).fill(0)]);
 const eof = rec(0xa,[]);
 const rk = (n: number) => rec(0x27e,[0,0,0,0,15,0,...u32((n << 2)|2)]);
 const bound = (offset: number, kind: number, name: string) => rec(0x85,[...u32(offset),0,kind,1,0,name.charCodeAt(0)]);
 const chart = [...bof(0x20),...eof];
 const extras = duplicate === undefined ? [] : duplicate === 0x27e ? rk(30) : rec(duplicate,[0,0,0,0,15,0,...new Array(14).fill(0)]);
 const a = [...bof(0x10),...rk(10),...extras,...eof];
 const b = [...bof(0x10),...rk(20),...eof];
 const prefix = [...bof(5),...bound(0,2,'C'),...bound(0,0,'B'),...bound(0,0,'A'),...eof];
 return new Uint8Array([...bof(5),...bound(prefix.length,2,'C'),...bound(prefix.length+chart.length+a.length,0,'B'),...bound(prefix.length+chart.length,0,'A'),...eof,...chart,...a,...b]);
}

describe('BIFF8 worksheet tab selection and ambiguous cells', () => {
 it('selects tab order through BOUNDSHEET offsets, excluding chart sheets', () => {
  const input = testBook();
  const result = editXlsNumericCell(input,{worksheetIndex:0,row:0,col:0,value:43});
  expect(result.status).toBe('edited');
  const before = readXlsWorkbook(input);
  const after = readXlsWorkbook(result.bytes);
  expect(before.sheets.map(s=>s.name)).toEqual(['C','B','A']);
  expect(after.sheets[1]?.cells[0]?.value).toBe(43);
  expect(after.sheets[2]?.cells[0]?.value).toBe(10);
  const next = editXlsNumericCell(result.bytes,{worksheetIndex:1,row:0,col:0,value:44});
  expect(next.status).toBe('edited');
  expect(readXlsWorkbook(next.bytes).sheets[2]?.cells[0]?.value).toBe(44);
  expect(editXlsNumericCell(input,{worksheetIndex:2,row:0,col:0,value:1})).toMatchObject({status:'unchanged',reason:'sheet-not-found'});
 });
 it.each(['middle-of-record','out-of-bounds','chart-pointer','duplicate-sheet'])('rejects invalid BOUNDSHEET mapping: %s', kind => {
  const input = testBook();
  const bounds = readRecords(input,0,input.length).filter(r=>r.opcode===0x85);
  const v = new DataView(input.buffer);
  const offset = kind==='middle-of-record' ? v.getUint32(bounds[1]!.dataOffset,true)+1 : kind==='out-of-bounds' ? input.length+4 : kind==='chart-pointer' ? v.getUint32(bounds[0]!.dataOffset,true) : v.getUint32(bounds[2]!.dataOffset,true);
  v.setUint32(bounds[1]!.dataOffset,offset,true);
  expect(editXlsNumericCell(input,{worksheetIndex:0,row:0,col:0,value:43})).toMatchObject({status:'unchanged',reason:'malformed-records',bytes:input});
 });
 it.each([0x27e,0x0006,0x00fd])('rejects duplicate target cell record %i', opcode => {
  const input = testBook(opcode);
  expect(editXlsNumericCell(input,{worksheetIndex:1,row:0,col:0,value:43})).toMatchObject({status:'unchanged',reason:'ambiguous-cell',bytes:input});
 });
});
