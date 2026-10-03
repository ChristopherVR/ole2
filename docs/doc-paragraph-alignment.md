# Direct legacy Word paragraph alignment

The typed model exposes a logical direct formatting exception on each stable
paragraph handle:

```ts
import { parseDoc } from '@christophervr/ole2';

const doc = parseDoc(bytes);
const paragraph = doc.paragraphs[1]!;
console.log(paragraph.directAlignment); // undefined when inherited or undecoded
paragraph.directAlignment = 'justify';
const saved = doc.serialize();
```

`DocParagraphAlignment` is `'start' | 'center' | 'end' | 'justify'`.
Start/end describe logical paragraph direction, without resolving inherited bidi
styles or claiming physical left/right positioning. The optional property on
the original `DocParagraph` interface retains compatibility with caller-created
paragraph objects; `ParsedDocParagraph` always has the accessor.

The setter replaces an existing unique `sprmPJc` (`0x2461`) operand in an
exclusive PAPX formatting run. Native Word also writes a preceding physical
`sprmPJc80` (`0x2403`) compatibility operand. When both unique values match, this
implementation updates both only for direction-independent `center` and
`justify`. Mirrored start/end writes throw `UnsupportedOle2EditError` with
`reason: 'direction-dependent-alignment'`; no inherited direction is guessed.
A sole modern slot supports all four logical values. Legacy-only, mismatched,
reversed, duplicate, missing or unrecognized formatting remains undefined or
refuses mutation. No formatting records are inserted.

See the official [MS-DOC paragraph properties](https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-doc/484822ee-a9d9-4af4-8423-29fda67a6a58)
and [PapxInFkp](https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-doc/580510b8-df7a-467e-a51c-0d71eb15c7cd)
definitions for physical/logical operand semantics and binary layout.

The writer checks piece PRMs, paragraph/run exclusivity, shared PAPX blobs,
other-story piece aliases, text and CHPX overlap, FIB-prefix overlap and section
exception overlap. Tables, fields/control characters in the target, external
huge PAPX and opaque/reset/style modifiers refuse. Default text/piece budgets
apply; paragraph formatting is bounded to 65,536 pages, physical runs and
paragraph records, and 262,144 SPRM records before allocation. Invalid values
are rejected without object coercion.

Paragraph handles read current bytes after text, run or alignment edits. A
successful edit increments the document revision and invalidates old character
run handles. Eligible no-op alignment assignments leave bytes, revision and
dirty unchanged. Refused edits commit nothing. Bytes outside the one modern
operand, or the two mirrored operands, remain identical.

## Owned native validation case

`test/fixtures/doc/paragraph-alignment.doc` is a neutral Word 16 derivative of
the owned rich-run fixture. Only the middle paragraph is directly centered;
neighbors inherit their alignment. Its generator and SHA-256 are retained in
the fixture provenance manifest. Native comparison covers center to justify
and exact center restoration, with captured paragraph formatting, text,
character fonts, styles, stories and counts checked independently. Sole-modern
start/end cases are synthetic codec regressions, not native fidelity evidence.
These checks do not establish rendering, pagination, every font/style, or full
legacy DOC fidelity.

After building, reproduce the consumer inputs without launching Word:

```sh
node scripts/prepare-native-doc-alignment-edits.mjs
```

The generator pins the source hash, checks exact changed-byte counts and emits
an explicit `nativeValidatedByThisRun: false` manifest. Native Word opening,
saving and reopening belongs to the independent corpus validation workflow.
