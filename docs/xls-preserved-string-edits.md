# Preservation-safe XLS string edits

`editXlsPreservedStringCell` replaces an existing BIFF8 `LABELSST` or `RK` cell
with plain Unicode text. It accepts a whole XLS compound file or a bare Workbook
stream. `worksheetIndex` follows worksheet tab order and excludes chart/macro
sheets; rows and columns are zero-based.

```ts
import { editXlsPreservedStringCell } from '@christophervr/ole2';

const result = editXlsPreservedStringCell(bytes, {
  worksheetIndex: 1,
  row: 0,
  col: 0,
  value: 'Updated 日本語 😀',
});
if (result.status === 'edited') {
  save(result.bytes);
  // Formula tokens and cached values are retained. Arrange recalculation in
  // the consuming spreadsheet application when dependent results must update.
} else {
  console.log(result.reason); // The exact caller bytes remain unchanged.
}
```

The writer appends a new SST entry or selects an existing plain entry. It retains
all previous SST/CONTINUE bytes, including rich runs and phonetic extension data,
so another cell sharing the old SST entry remains unchanged. The replacement is
plain text; replacing a rich target intentionally removes that target's inline
rich formatting while keeping its cell XF. Worksheet ROW, DBCELL, formula,
comment and other cell-table records are retained rather than rebuilt.

New strings can span CONTINUE records, up to 32,767 UTF-16 code units. BOUNDSHEET
and INDEX absolute offsets are relocated, and an existing ExtSST index is
regenerated. Unknown future record types and wrapped BIFF records are refused
when stream resizing would require an unimplemented pointer relocation model.
CFB resizing uses `resizeCompoundFileStream` and inherits its explicit layout
limits. All other stream payloads and directory metadata are preserved.

This increment does not add cells, convert NUMBER/MULRK records, write formulas,
calculate formulas, create workbooks, or support BIFF5/encryption. It rejects
malformed record framing, invalid pointers, duplicate target cells, incomplete
SST entries, malformed UTF-16, and SST reference counts inconsistent with the
existing LABELSST records. Every rejected operation returns the original input
reference and an explicit reason.

Tests use native Excel-generated public synthetic fixtures. Independent Excel
snapshots validate short Unicode and 24,000-unit continued Unicode edits, shared
aliases and selected later tabs. Raw-byte tests separately verify old rich SST
entries and untouched cell-table records. These checks establish the tested
editing surface; they do not claim complete spreadsheet fidelity.

The implementation follows [MS-XLS revision 12.2](https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-xls/),
particularly BoundSheet8, SST/Continue, XLUnicodeRichExtendedString, Index,
ExtSST and ISSTInf.
