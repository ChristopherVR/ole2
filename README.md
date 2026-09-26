# @christophervr/ole2

Framework-neutral primitives for legacy Office compound files. Shared by the PowerPoint and Word viewers without copying implementations.

```sh
npm install @christophervr/ole2
```

[Read the API guide and examples](https://christophervr.github.io/ole2/).

## Scope

- MS-CFB/OLE2 compound-file stream reading and writing, including mini streams, FAT/DIFAT and directory metadata.
- Word 97-2003 binary `.doc` main-body text reading and guarded existing-paragraph text editing.
- Excel BIFF8 `.xls` first-worksheet previews and bounded numeric/string cell edits.
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

## Development

```sh
bun install --frozen-lockfile
bun run typecheck
bun run test
bun run build
bun run test:package
```

The packed-package smoke test installs the tarball into an independent temporary npm project and exercises the public API. The build clears its own `dist` directory so removed codecs cannot accidentally remain in a release. Only compiled output, license, notice and readme ship to npm.
