# Legacy Excel

*XLS · BIFF8*

`parseXls` returns an `XlsDocument` with tab-order sheets and stable cell handles. Workbook styles, cached formulas and merges are read-only snapshots; supported cell value setters preserve existing record structure.

## Document model

Use sheets[index].cell(row, col).value with zero-based coordinates, then serialize(). Numeric and plain-string writes invoke bounded record editors.

- Existing NUMBER/RK/MULRK numeric values require exact original encoding.
- Existing LABELSST/RK/NUMBER/MULRK cells can become plain Unicode strings, including continued SST entries.
- Existing physical BLANK/MULBLANK cells can become number, plain string, boolean or error values; missing cells and merged followers still refuse.
- Packed siblings, selected XF and unrelated records are preserved; supported BOUNDSHEET/INDEX/DBCELL/ExtSST pointers are updated.
- Formula caches remain saved values; `recalculationRequired` signals that a consuming application must recalculate.
- Existing NUMBER/RK/MULRK cells may become boolean/error BOOLERR records; existing BOOLERR values can be replaced or converted to finite numbers/plain strings. Cell type distinguishes formula, number, string, boolean, error and blank.

## Explicit limits

Capabilities describe an editing surface, not guaranteed eligibility of each target.

- No missing-cell creation, formula writing or evaluation.
- Unsafe pointer-bearing records and allocation layouts are refused atomically.
- A selected rich string becomes plain while other shared aliases retain their formatting.
- Charts/drawings and general workbook construction are not modeled by this adapter.
- String-to-BOOLERR conversions remain unsupported.

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
