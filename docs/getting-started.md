# Getting started

`@christophervr/ole2` parses legacy Office compound files into typed, editable document classes and serializes them again. It is a codec library, not a renderer: it gives you the content and structure, and your application decides how to display it.

| Format | Extensions | Document class |
| --- | --- | --- |
| Word 97-2003 | `.doc` | `DocDocument` |
| Excel BIFF8 | `.xls` | `XlsDocument` |
| PowerPoint 97-2003 | `.ppt` | `PptDocument` |
| Visio (binary, version 11) | `.vsd` | `VsdDocument` |
| Any MS-CFB compound file | any | `CfbDocument` |

Publisher (`.pub`) files can be inspected but not parsed into a document model.

## Installation

::: code-group

```sh [npm]
npm install @christophervr/ole2
```

```sh [pnpm]
pnpm add @christophervr/ole2
```

```sh [yarn]
yarn add @christophervr/ole2
```

```sh [bun]
bun add @christophervr/ole2
```

:::

The package is ESM with TypeScript declarations, has no runtime dependencies and works in Node.js 22 or later and in browsers. Inputs are `Uint8Array` or `ArrayBuffer`.

## Quick start

`parseOle2` returns a typed editable document. `kind` narrows the default union; a generic format key requires a runtime expectation:

```ts
import { parseOle2, parseDoc, parseXls, parsePpt, parseVsd } from '@christophervr/ole2';

const document = parseOle2(bytes);
if (document.kind === 'xls') document.sheets[0]!.cell(1, 0).value = 'Updated';

const ppt = parseOle2<'ppt'>(pptBytes, { expect: 'ppt' });
ppt.slides[0]!.texts[0]!.text = 'Native title updated';
const saved = ppt.serialize();
```

`parseDoc`, `parseXls`, `parsePpt` and `parseVsd` return the concrete document classes directly.

- Supported setters commit transactionally. An unsupported edit throws `UnsupportedOle2EditError` without changing bytes or dirty state.
- Default parsing falls back to a CFB inspection view with explicit diagnostics when an Office model is unsupported or ambiguous. A checked `expect` throws instead.
- Adding `vsd` widened the default union, so exhaustive `switch` statements must handle `kind: 'vsd'`.

Read [Document models](./formats/models.md) for the full mutation and serialization contract.

## Containers

```ts
import { buildOle2, parseOle2 } from '@christophervr/ole2';

const bytes = buildOle2(new Map([['Example', new Uint8Array([1, 2, 3])]]));
const file = parseOle2(bytes, { expect: 'cfb' });
console.log(file.getStream('Example'));
```

## Granular imports

Subpaths expose the individual modules, for example `@christophervr/ole2/legacy-ppt-writer`. The package root re-exports the public API.

## Where to go next

- [Format guides](./formats/models.md) describe what each format supports.
- [API](./api.md) covers the lower-level functions that predate the document models.
- [Limitations](./limitations.md) lists what is not supported.
- [PPT export](./ppt-export.md) covers building a `.ppt` from a deck model.
