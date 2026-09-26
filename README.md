# @christophervr/ole2

Framework-neutral primitives for legacy Office compound files. Shared by the PowerPoint and Word viewers without copying implementations.

```sh
npm install @christophervr/ole2
```

## Scope

- MS-CFB/OLE2 compound-file stream reading and writing, including mini streams, FAT/DIFAT and directory metadata.
- Word 97-2003 binary `.doc` main-body text reading and guarded existing-paragraph text editing.
- No browser or framework dependency; typed-array/ArrayBuffer inputs and ESM JavaScript with TypeScript declarations.

**Modern DOCX does not belong here.** Its parser, document model, serializer and embedded-document adapter live in `@christophervr/docx-core`. PowerPoint consumes that Word package when editing embedded DOCX payloads.

```js
import { buildOle2, parseOle2 } from '@christophervr/ole2';
const bytes = buildOle2(new Map([['Example', new Uint8Array([1, 2, 3])]]));
const file = parseOle2(bytes);
console.log(file.getStream('Example'));
```

Granular subpaths are available for `ole2-parser-read`, `ole2-parser-write`, `ole2-parser-types`, and the `ole-document-doc-*` binary Word modules. Legacy Word editing is deliberately restricted; see source guards for unsupported character-position tables, encryption, formatting and document structures. This is not a complete legacy Office renderer.

## Development

```sh
bun install --frozen-lockfile
bun run typecheck
bun run test
bun run build
bun run test:package
```

The packed-package smoke test installs the tarball into an independent temporary npm project and exercises the public API. The build clears its own `dist` directory so removed codecs cannot accidentally remain in a release. Only compiled output, license, notice and readme ship to npm.

## Publishing

CI verifies every main push and pull request. `.github/workflows/release.yml` publishes tags such as `v0.1.0`, or an existing tag selected through workflow dispatch. The workflow verifies that the tag matches the manifest and checked-out commit, builds/tests the artifact, refuses registry errors, and treats an already-published version as a no-op.

For the initial publication, authenticate locally with `npm login` and run `npm publish --access public` after validation, or configure a repository `NPM_TOKEN` secret with publish permission. Then configure npm trusted publishing for GitHub owner **ChristopherVR**, repository **ole2**, workflow **release.yml**, with direct publishing allowed and no environment restriction. Subsequent tag releases use GitHub OIDC and provenance without a long-lived token. The workflow uses Node 24 and npm 11.

## Provenance

The CFB and binary Word implementations and reference fixtures were extracted from ChristopherVR/pptx-viewer, Apache-2.0. See `LICENSE`, `NOTICE` and source headers. Code changes belong here; consumer repositories retain compatibility re-exports and integration tests.
