import { readFileSync } from 'node:fs';
import { describe, expect, it } from 'vitest';
import { XlsDocument } from '../src/xls-document.js';
import { UnsupportedOle2EditError } from '../src/ole2-document-base.js';
import { readXlsWorkbook } from '../src/legacy-excel-workbook.js';
import { editXlsPreservedStringCell } from '../src/legacy-excel-preserved-string-cell.js';
import { buildOle2 } from '../src/ole2-parser-write.js';
const fixture = (name = 'workbook-features.xls') => new Uint8Array(readFileSync(new URL(`./fixtures/xls/${name}`, import.meta.url)));

describe('XlsDocument transactional adapter', () => {
 it('reads structured tab/cell/style/formula/merge metadata and serializes byte-identically', () => {
  const bytes = fixture(), document = new XlsDocument(bytes);
  expect(document.kind).toBe('xls');
  expect(document.workbook).toEqual(readXlsWorkbook(bytes));
  expect(document.sheets.map(s => s.name)).toEqual(document.workbook.sheets.map(s => s.name));
  const cell = document.sheets[0]!.cell(1, 1);
  expect(cell.value).toBe(1.25);
  expect(cell.style).toEqual(document.workbook.xfs[cell.xf]);
  expect(document.sheets[0]!.merges).toEqual(document.workbook.sheets[0]!.merges);
  expect(document.serialize()).toEqual(bytes);
  expect(document.dirty).toBe(false);
  expect(document.recalculationRequired).toBe(false);
  bytes.fill(0);
  expect(document.sheets[0]!.cell(1, 1).value).toBe(1.25);
  const serialized = document.serialize(); serialized.fill(0);
  expect(document.serialize()[0]).toBe(0xd0);
 });
 it('commits numeric writes and keeps existing cell handles current across subsequent edits', () => {
  const document = new XlsDocument(fixture()), sheet = document.sheets[0]!, cell = sheet.cell(1, 1);
  cell.value = 2.75;
  expect(cell.value).toBe(2.75);
  expect(sheet.cell(1, 1).value).toBe(2.75);
  expect(document.dirty).toBe(true);
  expect(document.recalculationRequired).toBe(true);
  cell.value = 'Converted';
  expect(cell.value).toBe('Converted');
  expect(new XlsDocument(document.serialize()).sheets[0]!.cell(1, 1).value).toBe('Converted');
 });
 it.each([
  { name: 'workbook-features.xls', row: 1, col: 1 },
  { name: 'workbook-features.xls', row: 1, col: 2 },
  { name: 'workbook-features.xls', row: 4, col: 1 },
  { name: 'workbook-mulrk.xls', row: 0, col: 0 },
  { name: 'workbook-mulrk.xls', row: 0, col: 1 },
  { name: 'workbook-mulrk.xls', row: 0, col: 3 },
 ])('routes native-validated $name [$row,$col] conversions through the same checked writer', ({ name, row, col }) => {
  const input = fixture(name), document = new XlsDocument(input);
  const value = 'Adapter 日本語 😀';
  const expected = editXlsPreservedStringCell(input, { row, col, value });
  expect(expected.status).toBe('edited');
  document.sheets[0]!.cell(row, col).value = value;
  expect(document.serialize()).toEqual(expected.bytes);
  const before = readXlsWorkbook(input), after = readXlsWorkbook(document.serialize());
  after.sheets[0]!.cells.find(c => c.row === row && c.col === col)!.value = before.sheets[0]!.cells.find(c => c.row === row && c.col === col)!.value;
  expect(after).toEqual(before);
 });
 it('refuses missing cells, formulas, unsafe numeric values and type changes atomically', () => {
  const input = fixture(), document = new XlsDocument(input), sheet = document.sheets[0]!;
  expect(() => sheet.cell(65535, 255)).toThrow(UnsupportedOle2EditError);
  expect(() => sheet.cell(-1, 0)).toThrow(UnsupportedOle2EditError);
  const formula = sheet.cells.find(c => c.formula !== undefined || c.formulaUndecoded)!;
  expect(formula).toBeDefined();
  expect(() => { formula.value = 3; }).toThrow(UnsupportedOle2EditError);
  const numeric = sheet.cell(1, 1);
  expect(() => { numeric.value = Math.PI; }).toThrow(UnsupportedOle2EditError);
  expect(() => { numeric.value = Infinity; }).toThrow(UnsupportedOle2EditError);
  expect(() => { numeric.value = true; }).toThrow(UnsupportedOle2EditError);
  const text = sheet.cells.find(c => typeof c.value === 'string' && !c.formula)!;
  expect(() => { text.value = 10; }).toThrow(UnsupportedOle2EditError);
  expect(document.serialize()).toEqual(input);
  expect(document.dirty).toBe(false);
  expect(document.recalculationRequired).toBe(false);
 });
 it('isolates frozen model snapshots and treats same-value assignment as clean', () => {
  const document = new XlsDocument(fixture()), cell = document.sheets[0]!.cell(1, 1);
  const model = document.workbook;
  expect(() => { (model.sheets[0]!.cells[0] as { value: unknown }).value = 'tamper'; }).toThrow();
  (model.numberFormats as Map<number, string>).set(999, 'tamper');
  expect(document.workbook.numberFormats.has(999)).toBe(false);
  cell.value = cell.value;
  expect(document.dirty).toBe(false);
  expect('model' in document).toBe(false);
  expect('sheetNodes' in document).toBe(false);
  expect(() => { (document as unknown as { kind: string }).kind = 'doc'; }).toThrow();
  expect(() => { (document.capabilities.write as string[]).push('formula writes'); }).toThrow();
  expect(() => { (cell as unknown as { row: number }).row = 100; }).toThrow();
 });
 it('rejects another compound format and corrupt input instead of casting it to XLS', () => {
  expect(() => new XlsDocument(new Uint8Array(buildOle2(new Map([['WordDocument', new Uint8Array(4096)]]))))).toThrow();
  expect(() => new XlsDocument(new Uint8Array([1, 2, 3]))).toThrow();
 });
});
