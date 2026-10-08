import { readFileSync } from 'node:fs';
import { describe, expect, it } from 'vitest';
import { editXlsNumericCell } from '../src/legacy-excel-biff8-edit.js';
import { readRecords } from '../src/legacy-excel-biff8.js';
import { unwrapXlsBytes } from '../src/legacy-excel-cfb.js';
import { readXlsWorkbook } from '../src/legacy-excel-workbook.js';
import { parseXls } from '../src/ole2-document.js';
import { readCompoundFileStream } from '../src/ole2-stream-edit.js';

const fixture = (name = 'workbook-mulrk.xls') => new Uint8Array(readFileSync(new URL(`./fixtures/xls/${name}`, import.meta.url)));
const raw = (bytes: Uint8Array) => unwrapXlsBytes(bytes).workbookBytes;
const cell = (bytes: Uint8Array, row: number, col: number) => readXlsWorkbook(bytes).sheets[0]!.cells.find(c => c.row === row && c.col === col)!;

function pointers(bytes: Uint8Array): void {
 const stream = raw(bytes), records = readRecords(stream, 0, stream.length), view = new DataView(stream.buffer, stream.byteOffset, stream.byteLength);
 const byOffset = new Map(records.map(r => [r.headerOffset, r]));
 for (const r of records) {
  if (r.opcode === 0x85) expect(byOffset.get(view.getUint32(r.dataOffset, true))?.opcode).toBe(0x809);
  if (r.opcode === 0x20b) for (let p = r.dataOffset + 12; p < r.dataOffset + r.length; p += 4)
   expect(byOffset.get(view.getUint32(p, true))?.opcode).toBe(p === r.dataOffset + 12 ? 0x55 : 0xd7);
  if (r.opcode !== 0xd7 || !view.getUint32(r.dataOffset, true)) continue;
  const first = byOffset.get(r.headerOffset - view.getUint32(r.dataOffset, true))!;
  expect(first.opcode).toBe(0x208);
  const rows = records.filter(row => row.opcode === 0x208 && row.headerOffset >= first.headerOffset && row.headerOffset < r.headerOffset);
  expect(rows.length).toBe((r.length - 4) / 2);
  let base = first.dataOffset + first.length;
  rows.forEach((row, i) => {
   const address = base + view.getUint16(r.dataOffset + 4 + i * 2, true), target = byOffset.get(address)!;
   expect(target).toBeDefined();
   expect(view.getUint16(target.dataOffset, true)).toBe(view.getUint16(row.dataOffset, true));
   base = address;
  });
 }
}

describe('lossless packed numeric promotion on native Excel fixtures', () => {
 it('promotes an original standalone RK on a native worksheet', () => {
  const input = fixture('workbook-features.xls'), before = readXlsWorkbook(input);
  const result = editXlsNumericCell(input, {row: 2, col: 5, value: -Number.MIN_VALUE});
  expect(result.status).toBe('edited'); expect(Object.is(cell(result.bytes, 2, 5).value, -Number.MIN_VALUE)).toBe(true);
  const after = readXlsWorkbook(result.bytes); after.sheets[0]!.cells.find(c => c.row === 2 && c.col === 5)!.value = before.sheets[0]!.cells.find(c => c.row === 2 && c.col === 5)!.value;
  expect(after).toEqual(before); pointers(result.bytes);
 });
 it.each([0, 1, 2, 3])('promotes MULRK column %i to an exact NUMBER and preserves sibling raw RK bytes', col => {
  const input = fixture(), original = readXlsWorkbook(input), source = raw(input), sourceView = new DataView(source.buffer, source.byteOffset, source.byteLength);
  const packed = readRecords(source, 0, source.length).find(r => r.opcode === 0xbd && sourceView.getUint16(r.dataOffset, true) === 0)!;
  const result = editXlsNumericCell(input, {row: 0, col, value: Math.PI});
  expect(result.status).toBe('edited'); expect(cell(result.bytes, 0, col).value).toBe(Math.PI);
  const after = readXlsWorkbook(result.bytes); after.sheets[0]!.cells.find(c => c.row === 0 && c.col === col)!.value = original.sheets[0]!.cells.find(c => c.row === 0 && c.col === col)!.value;
  expect(after).toEqual(original); pointers(result.bytes);
  const next = raw(result.bytes), view = new DataView(next.buffer, next.byteOffset, next.byteLength);
  for (const r of readRecords(next, 0, next.length).filter(r => [0x27e, 0xbd].includes(r.opcode) && view.getUint16(r.dataOffset, true) === 0)) {
   const first = view.getUint16(r.dataOffset + 2, true), count = r.opcode === 0xbd ? (r.length - 6) / 6 : 1;
   for (let i = 0; i < count; i++) expect(next.subarray(r.dataOffset + 4 + i * 6, r.dataOffset + 10 + i * 6)).toEqual(source.subarray(packed.dataOffset + 4 + (first + i) * 6, packed.dataOffset + 10 + (first + i) * 6));
  }
  for (const name of ['\u0005SummaryInformation', '\u0005DocumentSummaryInformation']) expect(readCompoundFileStream(result.bytes, [name])).toEqual(readCompoundFileStream(input, [name]));
 });
 it('promotes a split single RK, then retains handles across mixed exact and promoted edits', () => {
  const document = parseXls(fixture()), handles = [0, 1, 2, 3].map(col => document.sheets[0]!.cell(0, col));
  handles[1]!.value = Math.PI; // Leaves column 0 as a single RK.
  handles[0]!.value = Number.MIN_VALUE;
  handles[2]!.value = -Math.E;
  handles[3]!.value = 12.5;
  expect(handles.map(h => h.value)).toEqual([Number.MIN_VALUE, Math.PI, -Math.E, 12.5]);
  pointers(document.serialize());
 });
 it('refuses unknown pointer records and malformed DBCELL atomically', () => {
  const source = raw(fixture()), records = readRecords(source, 0, source.length);
  const unknown = source.slice(), harmless = records.find(r => r.opcode === 0x208)!;
  new DataView(unknown.buffer).setUint16(harmless.headerOffset, 0x7777, true);
  const rejected = editXlsNumericCell(unknown, {row: 0, col: 1, value: Math.PI});
  expect(rejected).toMatchObject({status: 'unchanged', reason: 'unsupported-pointer-record'}); expect(rejected.bytes).toBe(unknown);
  const malformed = source.slice(), db = records.find(r => r.opcode === 0xd7 && new DataView(source.buffer, source.byteOffset, source.byteLength).getUint32(r.dataOffset, true) !== 0)!;
  new DataView(malformed.buffer).setUint32(db.dataOffset, 1, true);
  const bad = editXlsNumericCell(malformed, {row: 0, col: 1, value: Math.PI});
  expect(bad).toMatchObject({status: 'unchanged', reason: 'malformed-records'}); expect(bad.bytes).toBe(malformed);
 });
});

it('refuses promotion of a merged follower and shared formula member atomically', () => {
 const rec = (op: number, data: number[]) => [op & 255, op >>> 8, data.length & 255, data.length >>> 8, ...data];
 const bof = (kind: number) => rec(0x809, [0, 6, kind, 0, ...new Array(12).fill(0)]), eof = rec(0xa, []);
 const bound = (offset: number) => rec(0x85, [offset & 255, offset >>> 8, 0, 0, 0, 0, 1, 0, 65]);
 const prefix = [...bof(5), ...bound(0), ...eof];
 for (const guard of [rec(0xe5, [1, 0, 0, 0, 0, 0, 0, 0, 1, 0]), rec(0x4bc, [0, 0, 0, 0, 0, 1])]) {
  const input = new Uint8Array([...bof(5), ...bound(prefix.length), ...eof, ...bof(0x10), ...rec(0x27e, [0, 0, 1, 0, 15, 0, 42, 0, 0, 0]), ...guard, ...eof]);
  const result = editXlsNumericCell(input, {row: 0, col: 1, value: Math.PI});
  expect(result).toMatchObject({status: 'unchanged', reason: 'cell-not-supported'}); expect(result.bytes).toBe(input);
 }
});
