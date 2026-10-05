# API

The primary API is `parseOle2` and the concrete `parseDoc`, `parseXls`, `parsePpt` and `parseVsd` helpers. See [Document models](./formats/models.md). This page covers the lower-level functions, which remain available and compatible.

The `ole2-parser-*`, `ole-document-doc-*`, `legacy-excel-*` and `legacy-ppt-*` modules are importable as granular subpaths.

## Format detection and Excel previews

```js
import {
  identifyLegacyOffice,
  readOleXlsGrid,
  writeOleXlsNumericCellEdit,
} from '@christophervr/ole2';

const kind = identifyLegacyOffice(fileBytes); // Uint8Array or ArrayBuffer
const grid = readOleXlsGrid(xlsBytes); // Uint8Array; first worksheet, bounded preview
const edited = writeOleXlsNumericCellEdit(xlsBytes, { row: 0, col: 0, value: 42 });
if (edited === xlsBytes) console.log('This edit was not supported.');
```

Unsupported edits return the exact input byte array; check that result before reporting success.

## Excel

`readXlsWorkbook(xlsBytes)` returns a structured `XlsWorkbook`: sheets, cells, XF styles with fonts, fills, borders and number formats, merges, views, comments, hyperlinks and names. It throws `XlsReadError` with `code` `'encrypted'`, `'unsupported-version'` (BIFF5 and earlier) or `'corrupt'`.

It never evaluates formulas: cells carry the cached result Excel saved, and the formula text when every token is understood. Charts, pictures, shapes, conditional formatting, data validation, pivot tables and VBA are listed in `unsupported` rather than decoded.

### Numeric edits

```ts
const result = editXlsNumericCell(bytes, { worksheetIndex: 0, row: 1, col: 1, value: 2.75 });
```

Supports existing NUMBER, RK and MULRK cells and returns `status: 'edited'` with `bytes` and `recalculationRequired: true`, or `status: 'unchanged'` with the original bytes and a reason. RK values must be exactly representable. This fixed-length path preserves all other records and compound-file bytes, including nested streams. Formula tokens and saved cached results are retained; nothing is recalculated.

### String edits

`editXlsStringCell` returns an explicit outcome for the existing first-worksheet string editor. Supported string changes retain existing rich shared-string data and UTF-16 characters.

`editXlsPreservedStringCell` is a compatibility editor for existing LABELSST/RK/NUMBER/MULRK cells. It retains original SST/CONTINUE data, neighboring cell records and selected XF, and updates supported workbook pointers. New application code can assign `parseXls(bytes).sheets[0].cell(row, col).value` and call `serialize()`. See [the preservation API guide](./xls-preserved-string-edits.md).

## Metadata

```js
import { readLegacyOfficeMetadata, writeLegacyOfficeMetadata } from '@christophervr/ole2';

console.log(readLegacyOfficeMetadata(fileBytes));
const edited = writeLegacyOfficeMetadata(fileBytes, 'title', 'New title');
if (edited === fileBytes) console.log('Unchanged or unsupported edit.');
```

Edits target an existing SummaryInformation property and the value must fit its allocated slot. Supported string encodings are Windows-1250 through Windows-1258 and UTF-16.

## Word

`tryWriteOleDocParagraphEdit(bytes, paragraphIndex, text)` returns either `status: 'edited'` with its strategy or `status: 'rejected'` with the original bytes and a reason. Equal-length plain-text edits retain original runs and character-position tables, including untouched headers and fields elsewhere. Growing edits retain paragraph and first-run formatting.

DOC readers and editors accept optional `OleDocProcessingLimits`. Defaults bound main text to 16,777,216 UTF-16 units and piece/field processing to 65,536 entries before allocation. The document model uses these defaults.

## PowerPoint

`readPptSlideTexts(bytes)` follows the active user-edit/persist directory rather than scanning stale saves. It returns slide ids and outline/inline text atoms. Encrypted input is rejected with `PptTextError.code === 'encrypted'`.

```js
import { readPptSlideTexts, editPptSlideText } from '@christophervr/ole2';

const atom = readPptSlideTexts(pptBytes).slides[0].texts[1];
const result = editPptSlideText(pptBytes, {
  slideIndex: 0, textIndex: 1,
  expectedText: atom.text, text: replacement,
});
if (result.status === 'edited') save(result.bytes);
else if (result.status === 'unsupported') console.log(result.reason);
```

The replacement must fit the existing encoding and keep its UTF-16 character count. Paragraph and control characters and field markers must remain at the same offsets. Only the selected text payload changes; unknown records, nested streams, images and save history stay byte-for-byte intact. Outline and inline text are listed separately by storage order.

To create a new deck, see [PPT export](./ppt-export.md).

## Container resizing

```ts
const result = resizeCompoundFileStream(bytes, ['ObjectPool', 'Workbook'], replacement);
```

Success returns `{ ok: true, bytes }`; refusal returns `{ ok: false, bytes: originalInput, reason }`. Mini streams may grow, shrink, become empty or move to and from regular allocations, and existing mini IDs, root bytes and directory metadata survive root growth. Allocations are appended, so a smaller payload does not compact the physical file. Format-specific callers must update internal Office record offsets and lengths themselves. See [Limitations](./limitations.md#cfb-containers) for the supported layouts.
