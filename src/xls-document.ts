/** Structured BIFF8 document adapter over the preservation-safe cell editors. */
import { Ole2DocumentBase, Ole2DocumentError, UnsupportedOle2EditError } from './ole2-document-base.js';
import { readCompoundFileStream } from './ole2-stream-edit.js';
import { readXlsWorkbook, type XlsWorkbook, type XlsSheet, type XlsCell, type XlsCellValue } from './legacy-excel-workbook.js';
import { editXlsNumericCell } from './legacy-excel-biff8-edit.js';
import { editXlsPreservedStringCell } from './legacy-excel-preserved-string-cell.js';
import { editXlsBoolErrorCell, isXlsErrorValue, normalizeXlsErrorValue } from './legacy-excel-bool-error-edit.js';
import { editXlsBlankCell } from './legacy-excel-blank-edit.js';

export type XlsCellType = 'number' | 'string' | 'boolean' | 'error' | 'blank' | 'formula';
function cellType(cell: XlsCell): XlsCellType {
 if (cell.formula !== undefined || cell.formulaUndecoded || cell.arrayRange || cell.sharedFormula) return 'formula';
 if (cell.value === null) return 'blank';
 if (typeof cell.value === 'object') return 'error';
 return typeof cell.value as 'number' | 'string' | 'boolean';
}
function sameValue(left: XlsCellValue, right: XlsCellValue): boolean {
 return Object.is(left, right) || (isXlsErrorValue(left) && isXlsErrorValue(right) && left.error === right.error);
}

export type XlsReadonly<T> = T extends Map<infer K, infer V> ? ReadonlyMap<K, XlsReadonly<V>>
 : T extends readonly (infer V)[] ? readonly XlsReadonly<V>[]
 : T extends object ? { readonly [K in keyof T]: XlsReadonly<T[K]> } : T;

function snapshot<T>(value: T): XlsReadonly<T> {
 const copy = structuredClone(value);
 const freeze = (item: unknown): void => {
  if (!item || typeof item !== 'object') return;
  if (item instanceof Map) for (const entry of item.values()) freeze(entry);
  else for (const entry of Object.values(item)) freeze(entry);
  Object.freeze(item);
 };
 freeze(copy);
 return copy as XlsReadonly<T>;
}

/** Existing cell handle. Only value writes use binary editors; metadata is a snapshot. */
export class XlsCellNode {
 constructor(private readonly document: XlsDocument, readonly sheetIndex: number, readonly row: number, readonly col: number) { Object.freeze(this); }
 get value(): XlsCellValue { return structuredClone(this.document.cellModel(this.sheetIndex, this.row, this.col).value); }
 set value(value: XlsCellValue) { this.document.setCellValue(this.sheetIndex, this.row, this.col, value); }
 get snapshot(): XlsReadonly<XlsCell & { type: XlsCellType }> {
  const cell = this.document.cellModel(this.sheetIndex, this.row, this.col);
  return snapshot({ ...cell, type: cellType(cell) });
 }
 get xf(): number { return this.snapshot.xf; }
 get type(): XlsCellType { return cellType(this.document.cellModel(this.sheetIndex, this.row, this.col)); }
 get formula(): string | undefined { return this.snapshot.formula; }
 get formulaUndecoded(): boolean | undefined { return this.snapshot.formulaUndecoded; }
 get style() { return this.document.workbook.xfs[this.xf]; }
}

/** Tab-order sheet handle, including read-only chart/macro tabs. */
export class XlsSheetNode {
 constructor(private readonly document: XlsDocument, readonly index: number) { Object.freeze(this); }
 get snapshot(): XlsReadonly<XlsSheet> { return this.document.workbook.sheets[this.index]!; }
 get name(): string { return this.snapshot.name; }
 get kind() { return this.snapshot.kind; }
 get state() { return this.snapshot.state; }
 get merges() { return this.snapshot.merges; }
 get cells(): readonly XlsCellNode[] { return Object.freeze(this.snapshot.cells.map(c => new XlsCellNode(this.document, this.index, c.row, c.col))); }
 cell(row: number, col: number): XlsCellNode {
  this.document.cellModel(this.index, row, col);
  return new XlsCellNode(this.document, this.index, row, col);
 }
}

const XLS_CAPABILITIES = Object.freeze({
  read: Object.freeze(['BIFF8 worksheets', 'cell values and cached formulas', 'styles and merges', 'opaque compound streams']),
  write: Object.freeze(['existing NUMBER/RK/MULRK numeric values', 'existing LABELSST/NUMBER/RK/MULRK plain string values', 'existing NUMBER/RK/MULRK/BOOLERR boolean and error values', 'existing BLANK/MULBLANK cells to number, plain string, boolean or error values']),
  limitations: Object.freeze(['No implicit cell creation, formula editing or recalculation', 'Numeric replacement of existing numeric cells requires exact original encoding', 'Resizing writes reject unsupported relocation records and container layouts', 'Plain string replacement removes selected rich text runs', 'String-to-boolean/error and boolean/error-to-number/string conversions unsupported']),
 });
/** BIFF8 Excel document. Does not create cells or recalculate formula caches. */
export class XlsDocument extends Ole2DocumentBase {
 get kind(): 'xls' { return 'xls'; }
 get capabilities() { return XLS_CAPABILITIES; }
 readonly #sheetNodes: readonly XlsSheetNode[];
 get sheets(): readonly XlsSheetNode[] { return this.#sheetNodes; }
 #model: XlsWorkbook;
 #needsRecalculation = false;
 #settingCellValue = false;
 constructor(input: Uint8Array) {
  super(input);
  const bytes = this.getBytes();
  const streams = ['Workbook', 'Book'].filter(name => readCompoundFileStream(bytes, [name]) !== undefined);
  if (streams.length === 0) throw new Ole2DocumentError('format-mismatch', 'The compound file has no root Excel workbook stream.');
  if (streams.length !== 1) throw new Ole2DocumentError('ambiguous-format', 'The compound file has multiple Excel workbook streams.');
  this.#model = readXlsWorkbook(bytes);
  this.#sheetNodes = Object.freeze(this.#model.sheets.map((_, index) => new XlsSheetNode(this, index)));
 }
 get workbook(): XlsReadonly<XlsWorkbook> { return snapshot(this.#model); }
 get recalculationRequired(): boolean { return this.#needsRecalculation; }
 /** @internal Used by stable cell handles; returns an isolated read snapshot. */
 cellModel(sheetIndex: number, row: number, col: number): XlsCell {
  if (!Number.isInteger(row) || row < 0 || row > 65535 || !Number.isInteger(col) || col < 0 || col > 255)
   throw new UnsupportedOle2EditError('invalid-cell-coordinate');
  const cells = this.#model.sheets[sheetIndex]?.cells.filter(c => c.row === row && c.col === col);
  if (!cells?.length) throw new UnsupportedOle2EditError('cell-not-found');
  if (cells.length !== 1) throw new UnsupportedOle2EditError('ambiguous-cell');
  return structuredClone(cells[0]!);
 }
 /** @internal All changes are validated and parsed before committing bytes. */
 setCellValue(sheetIndex: number, row: number, col: number, value: XlsCellValue): void {
  if (this.#settingCellValue) throw new UnsupportedOle2EditError('reentrant-cell-edit');
  this.#settingCellValue = true;
  try { this.#applyCellValue(sheetIndex, row, col, value); }
  finally { this.#settingCellValue = false; }
 }
 #applyCellValue(sheetIndex: number, row: number, col: number, value: XlsCellValue): void {
  if (value !== null && typeof value === 'object') {
   const normalized = normalizeXlsErrorValue(value);
   if (!normalized) throw new UnsupportedOle2EditError('value-type-not-supported');
   value = normalized;
  }
  const cell = this.cellModel(sheetIndex, row, col);
  if (cell.formula !== undefined || cell.formulaUndecoded || cell.arrayRange || cell.sharedFormula)
   throw new UnsupportedOle2EditError('formula-edit-not-supported');
  const sheet = this.#model.sheets[sheetIndex];
  if (sheet?.kind !== 'worksheet') throw new UnsupportedOle2EditError('sheet-not-supported');
  if (cell.value === null && value === null) return;
  if (typeof value !== 'number' && typeof value !== 'string' && typeof value !== 'boolean' && !isXlsErrorValue(value))
   throw new UnsupportedOle2EditError('value-type-not-supported');
  if (typeof value === 'number' && ((cell.value !== null && typeof cell.value !== 'number') || !Number.isFinite(value)))
   throw new UnsupportedOle2EditError('numeric-type-change-not-supported');
  if (typeof value === 'string' && cell.value !== null && typeof cell.value !== 'number' && typeof cell.value !== 'string')
   throw new UnsupportedOle2EditError('string-type-change-not-supported');
  if ((typeof value === 'boolean' || isXlsErrorValue(value)) && typeof cell.value === 'string')
   throw new UnsupportedOle2EditError('bool-error-type-change-not-supported');
  if (sameValue(cell.value, value)) return;
  const worksheetIndex = this.#model.sheets.slice(0, sheetIndex).filter(s => s.kind === 'worksheet').length;
  const edit = { row, col, value, worksheetIndex };
  const result = cell.value === null ? editXlsBlankCell(this.getBytes(), { ...edit, value }) : typeof value === 'number'
   ? editXlsNumericCell(this.getBytes(), { ...edit, value })
   : typeof value === 'string' ? editXlsPreservedStringCell(this.getBytes(), { ...edit, value })
   : editXlsBoolErrorCell(this.getBytes(), { ...edit, value });
  if (result.status !== 'edited') throw new UnsupportedOle2EditError(result.reason);
  const next = readXlsWorkbook(result.bytes);
  const selected = next.sheets[sheetIndex]?.cells.filter(c => c.row === row && c.col === col);
  if (selected?.length !== 1 || !sameValue(selected[0]!.value, value))
   throw new UnsupportedOle2EditError('edit-verification-failed');
  this.commitBytes(result.bytes);
  this.#model = next;
  this.#needsRecalculation = true;
 }
}
