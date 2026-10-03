# @christophervr/ole2

[![npm version](https://img.shields.io/npm/v/%40christophervr%2Fole2.svg)](https://www.npmjs.com/package/@christophervr/ole2)
[![license](https://img.shields.io/npm/l/%40christophervr%2Fole2.svg)](https://github.com/ChristopherVR/ole2/blob/main/LICENSE)
[![types](https://img.shields.io/npm/types/%40christophervr%2Fole2.svg)](https://www.npmjs.com/package/@christophervr/ole2)

> Framework-neutral codecs for legacy Office compound files.

[API guide](https://christophervr.github.io/ole2/) | [npm](https://www.npmjs.com/package/@christophervr/ole2) | [Full docs](https://christophervr.github.io/ole2/) | [Source](https://github.com/ChristopherVR/ole2)

Framework-neutral primitives for legacy Office compound files. Shared by the PowerPoint and Word viewers without copying implementations.

## Install

```bash
npm install @christophervr/ole2
```

[Read the API guide and examples](https://christophervr.github.io/ole2/).

## Features

- Unified checked parsing into editable DOC/XLS/PPT/VSD classes and an explicit CFB container view; supported model setters preserve opaque bytes and serialize transactionally. Adding VSD widens the default union, so exhaustive switches must handle `kind: 'vsd'`.
- MS-CFB/OLE2 compound-file stream reading and writing, including mini streams, FAT/DIFAT and directory metadata.
- Word 97-2003 binary `.doc` main-body paragraphs and rich-run reading, supported exclusive direct bold/italic, font-size, underline and paragraph-alignment changes, and guarded existing-paragraph edits, including fixed-length edits that retain original formatting runs and character-position tables.
- Excel BIFF8 `.xls` workbook reading (`readXlsWorkbook`: every sheet, cell values, cached formula results and decoded formula text, styles, merges, column and row sizes, views, comments, hyperlinks, defined names), plus typed sheet/cell handles and bounded numeric/string/boolean/error edits, including existing physical blank cells and boolean/error-to-number/plain-string conversions.
- PowerPoint 97-2003 `.ppt` active-slide text, shape and speaker-notes reading, fixed-length slide and validated notes-body text edits, and exclusive small-anchor rectangle/textbox geometry edits that preserve surrounding bytes, plus binary export from a framework-neutral model (text, shapes, pictures, notes, embedded objects and optional RC4 encryption).
- Binary Visio version 11 pages/shapes, explicit stored transforms, UTF-16 text and move/line geometry. Same-length text uses bounded compression preserving both original allocation and exact decoded length; supported literal-transform edits also retain allocation offsets. Insufficient capacity or work budgets refuse. Earlier versions, general styles/masters/ShapeSheet evaluation, rich text and general native Visio fidelity remain unsupported. Publisher remains structural inspection.
- Path-based compound stream edits preserve nested storage layout and all bytes outside the edited stream.
- `resizeCompoundFileStream` grows or shrinks regular and mini streams in supported v3 containers, including transitions across the 4096-byte cutoff, while retaining nested storage paths, unknown streams and original directory metadata. It returns an explicit result and refusal reason for unsupported layouts.
- No browser or framework dependency; typed-array/ArrayBuffer inputs and ESM JavaScript with TypeScript declarations.

## Quick start

The primary API returns a typed editable document. `kind` narrows the default
union; a generic format key requires a runtime expectation:

```ts
import { parseOle2, parseDoc, parseXls, parsePpt, parseVsd } from '@christophervr/ole2';

const document = parseOle2(bytes);
if (document.kind === 'xls') document.sheets[0]!.cell(1, 0).value = 'Updated';
const ppt = parseOle2<'ppt'>(pptBytes, { expect: 'ppt' });
ppt.slides[0]!.texts[0]!.text = 'Native title updated';
const saved = ppt.serialize();
```

`parseDoc`, `parseXls` and `parsePpt` return concrete document classes. Supported
setters commit transactionally; unsupported edits throw without changing bytes
or dirty state. Default parsing falls back to a CFB inspection view with explicit
diagnostics when an Office model is unsupported or ambiguous. Existing stream
inspection and operation APIs remain compatible. See [document models and the
serialization contract](docs/document-model.md).

For a container-level resize, call `resizeCompoundFileStream(bytes, ['ObjectPool', 'Workbook'], replacement)`. Success returns `{ ok: true, bytes }`; refusal returns `{ ok: false, bytes: originalInput, reason }`. Supported v3 containers use 512-byte sectors and at most 109 header-listed FAT sectors. Mini streams may grow, shrink, become empty or move to/from regular allocations. Existing mini IDs, root bytes and directory metadata survive root growth. V4 regular-stream resizing with 4096-byte sectors is also supported; V4 mini-stream resizing/transitions and external DIFAT expansion remain explicitly unsupported. Allocations are appended, so a smaller payload does not compact the physical file; emptying a stream needs no new allocation. Format-specific callers must update internal Office record offsets and lengths themselves.

```js
import { buildOle2, parseOle2 } from "@christophervr/ole2";
const bytes = buildOle2(new Map([["Example", new Uint8Array([1, 2, 3])]]));
const file = parseOle2(bytes);
console.log(file.getStream("Example"));
```

## API

Granular subpaths expose the `ole2-parser-*`, `ole-document-doc-*`, `legacy-excel-*` and `legacy-ppt-*` modules. This is a codec library, not a complete legacy Office renderer. The API guide describes each format's supported subset.

```js
import {
  identifyLegacyOffice,
  readOleXlsGrid,
  writeOleXlsNumericCellEdit,
} from "@christophervr/ole2";

const kind = identifyLegacyOffice(fileBytes); // Uint8Array or ArrayBuffer
const grid = readOleXlsGrid(xlsBytes); // Uint8Array; first worksheet, bounded preview
const edited = writeOleXlsNumericCellEdit(xlsBytes, { row: 0, col: 0, value: 42 });
if (edited === xlsBytes) console.log("This edit was not supported.");
```

`readXlsWorkbook(xlsBytes)` returns a structured `XlsWorkbook` (sheets, cells, XF styles with fonts, fills, borders and number formats, merges, views, comments, hyperlinks, names). It throws `XlsReadError` with `code` `'encrypted'`, `'unsupported-version'` (BIFF5 and earlier) or `'corrupt'`. It never evaluates formulas: cells carry the cached result Excel saved, and the formula text when every token is understood. Charts, pictures, shapes, conditional formatting, data validation, pivot tables and VBA are listed in `unsupported` rather than decoded.

Excel previews do not evaluate formulas or resolve every continued shared string. Editing is constrained by record type, workbook structure and string-table layout. Unsupported edits return the exact input byte array; callers should check that result before reporting success.

For an explicit numeric edit outcome, use `editXlsNumericCell(bytes, { worksheetIndex: 0, row: 1, col: 1, value: 2.75 })`. It supports existing NUMBER, RK and MULRK cells and returns `status: 'edited'` with `bytes` and `recalculationRequired: true`, or `status: 'unchanged'` with the original bytes and a reason. RK values must be exactly representable; unsupported precision is refused. This fixed-length path preserves all other records and compound-file bytes, including nested streams. Formula tokens and saved cached results are retained; the library does not recalculate them.

`editXlsStringCell` provides an explicit outcome for the existing first-worksheet string editor. Continued SSTs, INDEX pointers, unknown cell-region records and complex resizing layouts are rejected. Supported string changes retain existing rich shared-string data and UTF-16 characters. This guarded path is not a general workbook writer.

`editXlsPreservedStringCell` remains a compatibility editor for existing LABELSST/RK/NUMBER/MULRK cells. It retains original SST/CONTINUE data, neighboring cell records and selected XF, updating supported workbook pointers. New application code can assign `parseXls(bytes).sheets[0].cell(row, col).value` and call `serialize()`. Selected rich text becomes plain; other aliases retain their original data. See [the preservation API guide](docs/xls-preserved-string-edits.md).

```js
import { readLegacyOfficeMetadata, writeLegacyOfficeMetadata } from "@christophervr/ole2";

console.log(readLegacyOfficeMetadata(fileBytes));
const edited = writeLegacyOfficeMetadata(fileBytes, "title", "New title");
// Existing SummaryInformation property only; the value must fit its allocated slot.
// Supported string encodings are Windows-1250 through Windows-1258 and UTF-16.
if (edited === fileBytes) console.log("Unchanged or unsupported edit.");
```

For Word, `tryWriteOleDocParagraphEdit(bytes, paragraphIndex, text)` returns either `status: 'edited'` with its strategy or `status: 'rejected'` with the original bytes and a reason. Equal-length plain-text edits retain original runs and character-position tables, including untouched headers and fields elsewhere. Growing edits require a simpler supported document and retain paragraph/first-run formatting. Supported v3 regular-stream containers retain nested storages, unknown streams and directory metadata; dependent character-position tables and unsupported allocation layouts are refused. This remains a paragraph editor rather than a complete Word document model or creator.

DOC codec readers/editors accept optional `OleDocProcessingLimits`. Defaults bound main text to 16,777,216 UTF-16 units and piece/field processing to 65,536 entries before allocation. The document model uses these bounded defaults. Growing/shrinking plain paragraphs outside balanced main-story fields can preserve field codes, results and flags; edits inside fields and other unsupported CP-dependent structures remain refused.

## Legacy PowerPoint export

`readPptSlideTexts(bytes)` follows the active user-edit/persist directory rather than scanning stale saves. It returns slide ids and outline/inline text atoms. It does not decode formatting, layout, masters, notes, pictures or animations into an editable presentation model. Encrypted input is rejected with `PptTextError.code === 'encrypted'`.

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

The replacement must fit the existing encoding and keep its UTF-16 character count. Paragraph/control characters and field markers must remain at the same offsets. Existing formatting and hyperlink ranges retain their original character positions. Only the selected text payload changes; unknown records, nested streams, images and save history remain byte-for-byte intact. Slides with an OOXML `metroBlob` mirror are refused because modern PowerPoint may prefer that content over the binary fallback. Outline and inline text are listed separately by storage order; shape-to-outline references are not resolved. This constrained edit is separate from creating a new deck with `buildPptFile`, and is not a general presentation roundtrip.

```js
import { buildPptFile } from "@christophervr/ole2/legacy-ppt-writer";

const pptBytes = await buildPptFile({
  widthEmu: 9144000,
  heightEmu: 5143500,
  slides: [{ shapes: [] }],
  pictures: [],
});
// Save pptBytes as a .ppt file. Pass { password: '...' } as a second argument
// only when legacy RC4 output is required.
```

The writer consumes the typed `WDeck` model. Converting a PowerPoint viewer document into that model remains in `pptx-viewer-core`; it consumes this shared binary implementation.

### Known limitations of the `.ppt` writer

- **Several slide masters are not yet verified in PowerPoint.** `WDeck.masters` writes one `MainMaster` per entry and `WSlide.masterIndex` points each slide at its own. Masters take slide ids `0x80000000`, `0x80000001`, ... and each slide's `SlideAtom.masterIdRef` names its master by that id; the first id is COM-verified for single-master decks, while the consecutive ids for further masters follow `[MS-PPT]` but have not yet been checked by reopening in PowerPoint. A deck without `masters` still writes one master from `masterStyles` and `master`.
- **Picture formats.** BLIPs are `png`, `jpg`, `emf`, `wmf` and `dib`, matching PowerPoint's own 97-2003 save. GIF, TIFF and SVG sources must be converted by the caller first (`pptx-viewer-core` stores a GIF's first frame and a TIFF as a compressed PNG, and rasterises SVG).
- **Media.** Embedded sound is written as a playable WAV `SoundCollectionContainer`. The model has no video record, so video and non-WAV audio should be passed as their poster picture, which is what PowerPoint's own 97-2003 save produces.
- **Encryption.** Password-protected output uses RC4 CryptoAPI only; the pre-CryptoAPI Office 95 scheme is not written.
- **Modern-only content.** Charts, SmartArt, ink and 3D models have no binary record of their own; callers write a preview picture and can attach the element's OOXML as a `metroBlob` so PowerPoint 2007+ reopens it as the native object.

## Development

```sh
bun install --frozen-lockfile
bun run typecheck
bun run test
bun run build
bun run test:package
```

The packed-package smoke test installs the tarball into an independent temporary npm project and exercises the public API. The build clears its own `dist` directory so removed codecs cannot accidentally remain in a release. Only compiled output, license, notice and readme ship to npm.

The Pages build contains the format guides by default. To include the browser demo and generated samples in a local Pages build, set `OLE2_BUILD_PAGES_DEMO=1` before running `bun run build:pages`.

## Documentation

[API guide](https://christophervr.github.io/ole2/) | [Source](https://github.com/ChristopherVR/ole2)

## License

Apache-2.0.
