# Legacy Excel

*XLS · BIFF8*

`parseXls` returns an `XlsDocument` with tab-order sheets and stable cell handles. Workbook styles, cached formulas and merges are read-only snapshots; supported cell value setters preserve existing record structure.

## Document model

Use sheets[index].cell(row, col).value with zero-based coordinates, then serialize(). Numeric and plain-string writes invoke bounded record editors.

- Existing NUMBER/RK/MULRK cells accept finite numeric values without rounding. Values that do not fit RK promote the selected cell to NUMBER when BOUNDSHEET/INDEX/DBCELL relocation and container resizing are verified; packed siblings retain their raw RK and XF bytes.
- Existing LABELSST/RK/NUMBER/MULRK cells can become plain Unicode strings, including continued SST entries.
- Existing physical BLANK/MULBLANK cells can become number, plain string, boolean or error values; missing cells and merged followers still refuse.
- `sheet.createNumericCell(row, col, value, xf)` explicitly creates an absent numeric cell after a verified numeric predecessor in an existing nonempty ROW/DBCELL block. Supply an existing cell XF index; the coordinates must remain inside the saved DIMENSIONS. Existing physical cells are never overwritten by this operation.
- `sheet.createCell(row, col, value, xf)` uses the same bounded insertion path for finite numbers, plain Unicode strings, booleans and defined Excel error values. Strings may use SST continuations. Creation is staged and commits once; conversion or allocation refusal leaves no placeholder cell. Null/blank and formula creation remain unsupported.
- Packed siblings, selected XF and unrelated records are preserved; supported BOUNDSHEET/INDEX/DBCELL/ExtSST pointers are updated.
- Formula caches remain saved values; `recalculationRequired` signals that a consuming application must recalculate.
- Existing NUMBER/RK/MULRK cells may become boolean/error BOOLERR records; existing BOOLERR values can be replaced or converted to finite numbers/plain strings. Cell type distinguishes formula, number, string, boolean, error and blank.
- Existing LABELSST strings can become finite numbers, booleans or defined Excel errors. Only the selected SST reference is removed; original string entries, rich runs, phonetic data and shared aliases remain serialized. String-to-string replacement retains its existing behavior.

## Explicit limits

Capabilities describe an editing surface, not guaranteed eligibility of each target.

- No implicit cell creation, new rows, insertion before a row's first cell, formula writing or evaluation. Explicit creation also refuses merged ranges, shared/array formula ranges, protected sheets and unverified row-block layouts.
- Unsafe pointer-bearing records and allocation layouts are refused atomically.
- A selected rich string becomes plain while other shared aliases retain their formatting.
- Charts/drawings and general workbook construction are not modeled by this adapter.
- String scalar conversions require a checked LABELSST/SST layout; legacy LABEL records, formula string results and unsafe relocation records refuse.

## Compatibility APIs

`readXlsWorkbook`, previews and the numeric/string operation functions remain available. New code should navigate the document model.

## Example

```ts
import { parseXls } from '@christophervr/ole2';

const document = parseXls(xlsBytes);
const cell = document.sheets[0].cell(1, 0);
cell.value = 'Unicode Ω 日本';
const saved = document.serialize();
console.log(document.recalculationRequired);
```

Explicit creation needs a known eligible row and an existing cell-format index:

```ts
const created = document.sheets[0].createCell(2, 6, 'Unicode Ω 日本', 15);
console.log(created.type); // string
```

The example coordinates/style are fixture-specific, not universal defaults.
`createXlsCell` and `createXlsCellWorkbookStream` expose checked low-level
creation with `edited`/`unchanged` results. Refused creation returns the original
input object. The original numeric creation APIs remain available.

Scalar creation tests use Excel-generated fixtures to verify original cells,
formulas, styles, unrelated compound streams and relocated BIFF pointers. They
also cover continued strings, error values, mixed edits, v4 containers and staged
conversion refusals. Native Excel open/save/reopen validation for these new
creation paths has not been performed.

String-to-number/boolean/error conversion is covered by the rich-string fixture
and a native Excel open/save/reopen comparison of five declared edits. See
[string scalar conversions](../xls-string-scalar-conversions.md) for scope and
reproduction.

Packed numeric promotion is covered by Excel-generated fixtures and byte/pointer checks. Fresh native Excel save/reopen validation is still required; these checks do not establish full application parity. Minimal bare BIFF streams without a tab directory retain the `inexact-rk` refusal.
