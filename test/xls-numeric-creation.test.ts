import {readFileSync} from 'node:fs';
import {describe, expect, it} from 'vitest';
import {createXlsNumericCell, createXlsNumericCellWorkbookStream} from '../src/legacy-excel-create-numeric-cell.js';
import {readXlsWorkbook} from '../src/legacy-excel-workbook.js';
import {unwrapXlsBytes} from '../src/legacy-excel-cfb.js';
import {readRecords} from '../src/legacy-excel-biff8.js';
import {readCompoundFileStream} from '../src/ole2-stream-edit.js';
import {parseXls} from '../src/ole2-document.js';

const fixture = (name = 'workbook-mulrk.xls') => new Uint8Array(readFileSync(new URL(`./fixtures/xls/${name}`, import.meta.url)));
const raw = (input: Uint8Array) => unwrapXlsBytes(input).workbookBytes;
const edit = {row: 3, col: 2, value: Math.PI, xf: 62};

function checkPointers(input: Uint8Array): void {
 const bytes = raw(input), view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength), records = readRecords(bytes, 0, bytes.length), offsets = new Map(records.map(r => [r.headerOffset, r]));
 for (const r of records) {
  if (r.opcode === 0x85) expect(offsets.get(view.getUint32(r.dataOffset, true))?.opcode).toBe(0x809);
  if (r.opcode === 0x20b) for (let p = r.dataOffset + 12; p < r.dataOffset + r.length; p += 4) expect(offsets.get(view.getUint32(p, true))?.opcode).toBe(p === r.dataOffset + 12 ? 0x55 : 0xd7);
  if (r.opcode !== 0xd7 || !view.getUint32(r.dataOffset, true)) continue;
  const first = offsets.get(r.headerOffset - view.getUint32(r.dataOffset, true))!;
  expect(first.opcode).toBe(0x208);
  const rows = records.filter(row => row.opcode === 0x208 && row.headerOffset >= first.headerOffset && row.headerOffset < r.headerOffset);
  expect(rows.length).toBe((r.length - 4) / 2);
  let base = first.dataOffset + first.length;
  for (let i = 0; i < rows.length; i++) {
   base += view.getUint16(r.dataOffset + 4 + i * 2, true);
   expect(view.getUint16(offsets.get(base)!.dataOffset, true)).toBe(view.getUint16(rows[i]!.dataOffset, true));
  }
 }
}

describe('explicit bounded absent numeric cell creation', () => {
 it.each([
  ['workbook-mulrk.xls', 3, 2, 62], ['workbook-features.xls', 2, 6, 15], ['workbook-features.xls', 3, 3, 15],
 ] as const)('adds one NUMBER to native %s row %i column %i and preserves all existing cells', (name, row, col, xf) => {
  const input = fixture(name), before = readXlsWorkbook(input), source = raw(input);
  expect(before.sheets[0]!.cells.some(c => c.row === row && c.col === col)).toBe(false);
  const result = createXlsNumericCell(input, {row, col, value: Math.PI, xf}); expect(result.status).toBe('edited');
  const after = readXlsWorkbook(result.bytes), created = after.sheets[0]!.cells.find(c => c.row === row && c.col === col)!;
  expect(created).toMatchObject({row, col, value: Math.PI, xf});
  after.sheets[0]!.cells = after.sheets[0]!.cells.filter(c => c !== created); expect(after).toEqual(before);
  checkPointers(result.bytes);
  const next = raw(result.bytes), oldRecords = readRecords(source, 0, source.length), newRecords = readRecords(next, 0, next.length);
  // Every original cell record remains serialized exactly; only one new record is inserted.
  const ops = new Set([0x6,0x201,0x203,0x204,0x205,0x27e,0xfd,0xbd,0xbe,0xd6]);
  const selected = (bytes: Uint8Array, records: typeof oldRecords) => records.filter(r => ops.has(r.opcode)).map(r => bytes.slice(r.headerOffset, r.dataOffset + r.length));
  const original = selected(source, oldRecords), actual = selected(next, newRecords);
  const view = new DataView(next.buffer, next.byteOffset, next.byteLength);
  const addition = newRecords.filter(r => ops.has(r.opcode)).findIndex(r => r.opcode === 0x203 && view.getUint16(r.dataOffset, true) === row && view.getUint16(r.dataOffset + 2, true) === col);
  actual.splice(addition, 1); expect(actual).toEqual(original);
  for (const name of ['\u0005SummaryInformation', '\u0005DocumentSummaryInformation']) expect(readCompoundFileStream(result.bytes, [name])).toEqual(readCompoundFileStream(input, [name]));
 });
 it('retains signed zero and stable handles through creation and mixed writes', () => {
  const document = parseXls(fixture()), sheet = document.sheets[0]!, existing = sheet.cell(3, 1);
  const created = sheet.createNumericCell(3, 2, -0, 62);
  expect(Object.is(created.value, -0)).toBe(true); expect(sheet.cell(3, 2).xf).toBe(62);
  created.value = Math.PI; existing.value = Math.E;
  const another = sheet.createNumericCell(3, 3, Number.MIN_VALUE, 63);
  expect(another.value).toBe(Number.MIN_VALUE); expect(created.value).toBe(Math.PI); expect(existing.value).toBe(Math.E);
  checkPointers(document.serialize()); expect(document.recalculationRequired).toBe(true);
 });
 it.each([{row: 0, col: 0}, {row: 4, col: 4}])('never overwrites an existing cell at $row/$col', target => {
  const input = fixture(), result = createXlsNumericCell(input, {...edit, ...target});
  expect(result.status).toBe('unchanged'); expect(result.bytes).toBe(input);
 });
 it.each([{row: 5}, {row: 65535}, {col: 5}, {row: 1, col: 3}])('refuses unsupported rows/extents/non-numeric predecessors: %o', target => {
  const input = fixture(), result = createXlsNumericCell(input, {...edit, ...target});
  expect(result.status).toBe('unchanged'); expect(result.bytes).toBe(input);
 });
 it.each([{value: Infinity}, {value: NaN}, {xf: -1}, {xf: 0}, {xf: 65535}, {row: -1}, {col: 256}])('refuses invalid coordinates, nonfinite numbers and invalid/style XFs: %o', target => {
  const input = fixture(), result = createXlsNumericCell(input, {...edit, ...target});
  expect(result).toMatchObject({status: 'unchanged', reason: 'invalid-edit'}); expect(result.bytes).toBe(input);
 });
 it('refuses malformed pointers, unknown future records and bad container bytes atomically', () => {
  const source = raw(fixture()), records = readRecords(source, 0, source.length), db = records.find(r => r.opcode === 0xd7)!, row = records.find(r => r.opcode === 0x208)!;
  for (const [record, value, expected] of [[db, 1, 'creation-not-supported'], [row, 0x7777, 'creation-not-supported']] as const) {
   const input = source.slice(), view = new DataView(input.buffer);
   if (record === db) view.setUint32(db.dataOffset, value, true); else view.setUint16(row.headerOffset, value, true);
   const result = createXlsNumericCellWorkbookStream(input, edit); expect(result.status).toBe('unchanged'); expect(result.bytes).toBe(input);
  }
  const unknown = source.slice(), harmless = records.find(r => r.opcode === 0x23e)!;
  new DataView(unknown.buffer).setUint16(harmless.headerOffset, 0x7777, true);
  const result = createXlsNumericCellWorkbookStream(unknown, edit); expect(result).toMatchObject({status: 'unchanged', reason: 'unsupported-pointer-record'}); expect(result.bytes).toBe(unknown);
  const bad = fixture().slice(0, 100), failed = createXlsNumericCell(bad, edit); expect(failed.status).toBe('unchanged'); expect(failed.bytes).toBe(bad);
 });
 it('leaves model bytes and dirty state unchanged after an explicit creation refusal', () => {
  const input = fixture(), document = parseXls(input);
  expect(() => document.sheets[0]!.createNumericCell(65535, 2, 1, 62)).toThrow();
  expect(document.serialize()).toEqual(input); expect(document.dirty).toBe(false); expect(document.recalculationRequired).toBe(false);
 });
 it('preserves already-dirty mixed edits when a later creation fails', () => {
  const document = parseXls(fixture()), sheet = document.sheets[0]!, handle = sheet.createNumericCell(3, 2, Math.PI, 62);
  sheet.cell(3, 1).value = Math.E; const committed = document.serialize();
  expect(() => sheet.createNumericCell(3, 2, 9, 62)).toThrow();
  expect(document.serialize()).toEqual(committed); expect(handle.value).toBe(Math.PI); expect(document.dirty).toBe(true); expect(document.recalculationRequired).toBe(true);
 });
 it('rejects merged/shared ranges covering an absent target, including malformed ranges', () => {
  const source = raw(fixture()), records = readRecords(source, 0, source.length), merged = records.find(r => r.opcode === 0xe5)!;
  const nonTargetRow = records.find(r => r.opcode === 0x208)!;
  for (const malformed of [false, true]) for (const kind of ['merged', 'shared', 'array'] as const) {
   const input = source.slice(), view = new DataView(input.buffer);
   if (kind === 'merged') {
    view.setUint16(merged.dataOffset + 2, 3, true); view.setUint16(merged.dataOffset + 4, malformed ? 2 : 3, true);
    view.setUint16(merged.dataOffset + 6, 2, true); view.setUint16(merged.dataOffset + 8, 3, true);
   } else {
    view.setUint16(nonTargetRow.headerOffset, kind === 'shared' ? 0x4bc : 0x221, true);
    view.setUint16(nonTargetRow.dataOffset, 3, true); view.setUint16(nonTargetRow.dataOffset + 2, malformed ? 2 : 3, true);
    view.setUint8(nonTargetRow.dataOffset + 4, 2); view.setUint8(nonTargetRow.dataOffset + 5, 3);
   }
   const result = createXlsNumericCellWorkbookStream(input, edit);
   expect(result).toMatchObject({status: 'unchanged', reason: malformed ? 'malformed-records' : 'creation-not-supported'}); expect(result.bytes).toBe(input);
  }
 });
 it('rejects duplicate/descending cells, false ROW extents and malformed DBCELL cell pointers', () => {
  const source = raw(fixture()), view = new DataView(source.buffer, source.byteOffset, source.byteLength), records = readRecords(source, 0, source.length);
  const targetRow = records.find(r => r.opcode === 0x208 && view.getUint16(r.dataOffset, true) === 3)!;
  const number = records.find(r => r.opcode === 0x203 && view.getUint16(r.dataOffset, true) === 3)!;
  const firstCell = records.find(r => r.opcode === 0x6 && view.getUint16(r.dataOffset, true) === 3 && view.getUint16(r.dataOffset + 2, true) === 0)!;
  const db = records.find(r => r.opcode === 0xd7 && view.getUint32(r.dataOffset, true) !== 0)!;
  for (const mutation of [
   (v: DataView) => v.setUint16(number.dataOffset + 2, 0, true),
   (v: DataView) => v.setUint16(firstCell.dataOffset + 2, 4, true),
   (v: DataView) => v.setUint16(targetRow.dataOffset + 4, 1, true),
   (v: DataView) => v.setUint16(db.dataOffset + 4, 65535, true),
  ]) {
   const input = source.slice(); mutation(new DataView(input.buffer)); const result = createXlsNumericCellWorkbookStream(input, edit);
   expect(result.status).toBe('unchanged'); expect(result.bytes).toBe(input);
  }
 });
 it('refuses pathological record counts before allocating the checked model', () => {
  const input = new Uint8Array(100001 * 4), result = createXlsNumericCellWorkbookStream(input, edit);
  expect(result).toMatchObject({status: 'unchanged', reason: 'creation-not-supported'}); expect(result.bytes).toBe(input);
 });
});
