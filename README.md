# @christophervr/ole2

Framework-neutral primitives for legacy Office compound files. Shared by the PowerPoint and Word viewers without copying implementations.

```sh
npm install @christophervr/ole2
```

[Read the API guide and examples](https://christophervr.github.io/ole2/).

## Scope

- MS-CFB/OLE2 compound-file stream reading and writing, including mini streams, FAT/DIFAT and directory metadata.
- Word 97-2003 binary `.doc` main-body text reading and guarded existing-paragraph text editing.
- Excel BIFF8 `.xls` workbook reading (`readXlsWorkbook`: every sheet, cell values, cached formula results and decoded formula text, styles, merges, column and row sizes, views, comments, hyperlinks, defined names), plus first-worksheet previews and bounded numeric/string cell edits.
- PowerPoint `.ppt` binary export from a framework-neutral model, including text, shapes, pictures, notes, embedded objects and optional RC4 encryption, plus record headers, traversal and constants.
- Visio and Publisher binary structure inspection, plus standard OLE document-property reading and bounded text-property edits. Drawing and publication page content is not yet decoded or editable.
- Path-based compound stream edits preserve nested storage layout and all bytes outside the edited stream.
- No browser or framework dependency; typed-array/ArrayBuffer inputs and ESM JavaScript with TypeScript declarations.

```js
import { buildOle2, parseOle2 } from '@christophervr/ole2';
const bytes = buildOle2(new Map([['Example', new Uint8Array([1, 2, 3])]]));
const file = parseOle2(bytes);
console.log(file.getStream('Example'));
```

Granular subpaths expose the `ole2-parser-*`, `ole-document-doc-*`, `legacy-excel-*` and `legacy-ppt-*` modules. This is a codec library, not a complete legacy Office renderer. The API guide describes each format's supported subset.

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

`readXlsWorkbook(xlsBytes)` returns a structured `XlsWorkbook` (sheets, cells, XF styles with fonts, fills, borders and number formats, merges, views, comments, hyperlinks, names). It throws `XlsReadError` with `code` `'encrypted'`, `'unsupported-version'` (BIFF5 and earlier) or `'corrupt'`. It never evaluates formulas: cells carry the cached result Excel saved, and the formula text when every token is understood. Charts, pictures, shapes, conditional formatting, data validation, pivot tables and VBA are listed in `unsupported` rather than decoded.

Excel previews do not evaluate formulas or resolve every continued shared string. Editing is constrained by record type, workbook structure and string-table layout. Unsupported edits return the exact input byte array; callers should check that result before reporting success.

```js
import { readLegacyOfficeMetadata, writeLegacyOfficeMetadata } from '@christophervr/ole2';

console.log(readLegacyOfficeMetadata(fileBytes));
const edited = writeLegacyOfficeMetadata(fileBytes, 'title', 'New title');
// Existing SummaryInformation property only; the value must fit its allocated slot.
// Supported string encodings are Windows-1250 through Windows-1258 and UTF-16.
if (edited === fileBytes) console.log('Unchanged or unsupported edit.');
```

## Legacy PowerPoint export

```js
import { buildPptFile } from '@christophervr/ole2/legacy-ppt-writer';

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
