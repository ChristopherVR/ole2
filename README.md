# @christophervr/ole2

[![npm version](https://img.shields.io/npm/v/%40christophervr%2Fole2.svg)](https://www.npmjs.com/package/@christophervr/ole2)
[![CI](https://github.com/ChristopherVR/ole2/actions/workflows/ci.yml/badge.svg)](https://github.com/ChristopherVR/ole2/actions/workflows/ci.yml)
[![license](https://img.shields.io/npm/l/%40christophervr%2Fole2.svg)](https://github.com/ChristopherVR/ole2/blob/main/LICENSE)
[![types](https://img.shields.io/npm/types/%40christophervr%2Fole2.svg)](https://www.npmjs.com/package/@christophervr/ole2)

Framework-neutral TypeScript codecs for **MS-CFB / OLE2** compound files and the legacy binary Office formats built on them. Parse a file into a typed, editable document, change supported content, and serialize it again while preserving the bytes you did not touch.

| Format | Extension | Document class |
| --- | --- | --- |
| Word 97-2003 | `.doc` | `DocDocument` |
| Excel BIFF8 | `.xls` | `XlsDocument` |
| PowerPoint 97-2003 | `.ppt` | `PptDocument` |
| Visio (binary) | `.vsd` | `VsdDocument` |
| Any compound file | any | `CfbDocument` |

[Documentation](https://christophervr.github.io/ole2/) · [Limitations](https://christophervr.github.io/ole2/limitations) · [npm](https://www.npmjs.com/package/@christophervr/ole2)

---

## Install

```bash
npm install @christophervr/ole2
```

No runtime dependencies. ESM with TypeScript declarations, for Node.js 22+ and browsers.

## Quick start

```ts
import { parseOle2, parseDoc, parseXls, parsePpt, parseVsd } from '@christophervr/ole2';

const document = parseOle2(bytes); // CFB | DOC | XLS | PPT | VSD
if (document.kind === 'xls') document.sheets[0]!.cell(1, 0).value = 'Updated';

const ppt = parseOle2<'ppt'>(pptBytes, { expect: 'ppt' });
ppt.slides[0]!.texts[0]!.text = 'Native title updated';
const saved = ppt.serialize();
```

`parseDoc`, `parseXls`, `parsePpt` and `parseVsd` return the concrete classes. Supported setters commit transactionally; an unsupported edit throws without changing bytes or dirty state. Default parsing falls back to a CFB inspection view with diagnostics when an Office model is unsupported or ambiguous, while a checked `expect` throws instead.

```ts
import { buildOle2, parseOle2 } from '@christophervr/ole2';

const bytes = buildOle2(new Map([['Example', new Uint8Array([1, 2, 3])]]));
console.log(parseOle2(bytes, { expect: 'cfb' }).getStream('Example'));
```

## What is included

- **Document models** for DOC, XLS, PPT and VSD with checked parsing and transactional edits.
- **MS-CFB containers**: stream reading and writing, mini streams, FAT/DIFAT, in-place stream resizing and SummaryInformation metadata.
- **Word**: main-body paragraphs, rich runs, direct formatting and guarded text edits.
- **Excel**: whole-workbook reading (cells, cached formula results, styles, merges, comments, hyperlinks, names) and typed cell edits.
- **PowerPoint**: active slides, text, shapes and notes, bounded text and geometry edits, and a binary `.ppt` writer (`buildPptFile`).
- **Visio**: version 11 pages, shapes, drawing order, transforms, text and geometry, with preserving text and transform edits.
- **Publisher**: container inspection.

Each format has a supported subset, which is documented on the site rather than here: see the [format guides](https://christophervr.github.io/ole2/formats/models) and the [limitations page](https://christophervr.github.io/ole2/limitations).

## Documentation

The documentation site is built with VitePress from the `docs/` directory.

- [Getting started](https://christophervr.github.io/ole2/getting-started)
- [Document models](https://christophervr.github.io/ole2/formats/models)
- [API reference](https://christophervr.github.io/ole2/api)
- [PPT export](https://christophervr.github.io/ole2/ppt-export)
- [Limitations](https://christophervr.github.io/ole2/limitations)

Run it locally with `bun run docs:dev`.

## Development

```sh
bun install --frozen-lockfile
bun run typecheck
bun run test
bun run build
bun run test:package
bun run docs:build
```

The packed-package smoke test installs the tarball into an independent temporary npm project and exercises the public API. The build clears its own `dist` directory so removed codecs cannot remain in a release. Only compiled output, license, notice and this readme ship to npm.

## License

Apache-2.0.
