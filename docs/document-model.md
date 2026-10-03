# Typed editable legacy documents

The primary API parses a compound file into an editable document class. This
follows the viewer's document/content navigation pattern while retaining opaque
binary bytes rather than pretending the legacy formats have a complete model.

```ts
import {parseOle2, parseDoc, parseXls, parsePpt} from '@christophervr/ole2';

const document = parseOle2(bytes); // CfbDocument | DocDocument | XlsDocument | PptDocument
if (document.kind === 'xls') document.sheets[0]!.cell(1, 0).value = 'Updated';

const ppt = parseOle2<'ppt'>(pptBytes, {expect: 'ppt'});
ppt.slides[0]!.texts[0]!.text = 'Native title updated';
const savedPpt = ppt.serialize();

const doc = parseDoc(docBytes);
doc.paragraphs[0]!.text = 'A longer plain paragraph outside fields.';
const savedDoc = doc.serialize();

const xls = parseXls(xlsBytes);
xls.sheets[0]!.cell(1, 0).value = 'Unicode Ω 日本';
const savedXls = xls.serialize();
```

All four classes extend `Ole2DocumentBase`. DOC owns stable paragraph handles;
XLS owns stable tab/cell handles and exposes read-only workbook, style, formula
and merge snapshots; PPT owns active slides with stable text identities. CFB
exposes existing stream handles by full storage path: `cfb.stream(['Custom',
'Data']).bytes = replacement`. This operation edits the container and does not
repair Office record offsets. No new directory entries are created.

## Checked parsing

The generic is a format key (`'cfb' | 'doc' | 'xls' | 'ppt'`), not a document-class
cast. The overload requires `{expect: K}` for `parseOle2<K>` and validates that
expectation at runtime. Both `parseOle2<'ppt'>(bytes)` and an explicit `'ppt'`
generic combined with `{expect: 'xls'}` fail TypeScript compilation.
`parseDoc`, `parseXls` and `parsePpt` return concrete checked classes.

Default parsing returns a union discriminated by the immutable `kind` getter.
Ambiguous root format streams or a malformed/encrypted/unsupported Office model
return a `CfbDocument` with explicit `diagnostics` and, when uniquely recognized,
`detectedFormat`. This is container inspection, not successful Office decoding.
Checked Office parsing throws `Ole2DocumentError` instead of returning a false
model. `{expect: 'cfb'}` explicitly selects container inspection for any valid
compound file. Invalid CFB allocation/header parsing still throws.

VSD/PUB page models are not implemented; those files remain container views with
the existing structural/metadata inspection codecs available for compatibility.

## Mutation and serialization

`capabilities.read`, `.write` and `.limitations` describe the supported surface,
not guaranteed eligibility of a particular target. Setters invoke the bounded
format writers, validate the candidate model and commit only supported changes.
An unsupported mutation throws `UnsupportedOle2EditError` with its reason before
changing bytes, model values, `dirty` or `revision`. A no-op leaves them unchanged.
Success marks `dirty` and increments `revision`. Existing handles observe later
supported edits rather than retaining stale snapshots.

`serialize()` returns a new byte array retaining all opaque data supported by the
underlying preservation strategy. It does not reset dirty state, write a file or
reconstruct unsupported content. Input bytes, returned streams, entry CLSIDs and
serialization output cannot alias owned bytes. Inspection snapshots are detached;
only supported model setters edit the document. Protected base methods are
internal adapter hooks, not a supported external mutation API.

| Model | Supported mutations | Explicit limitations |
| --- | --- | --- |
| DOC | Existing plain paragraph text; guarded growth/shrink outside balanced main-story fields | No paragraph insertion/removal, run/table/object model writes, field-code/result editing or unsupported CP-table shifts; processing budgets apply |
| XLS | Existing numeric cells; existing LABELSST/RK/NUMBER/MULRK cells to plain strings, including continued Unicode SSTs | No cell creation, formulas or recalculation; unsafe relocation records/layouts refused; selected rich string becomes plain while retaining XF and other aliases |
| PPT | Existing active text atoms with fixed UTF-16 length and original encoding | No shape-to-text mapping, shape/run/notes edits or arbitrary text growth; mirrors, field/control changes and shared physical atoms refused |
| CFB | Supported existing regular/mini stream resizing with path/hierarchy preservation | No new directory entries, external DIFAT expansion or variable-length v4 mini transitions; container support is separate from Office fidelity |

## Compatibility and migration

The root `parseOle2` keeps the complete `Ole2File` inspection shape (`entries` and
`getStream`) while returning the document union. `parseCompoundFile` names the
original raw reader. The original `ole2-parser-read` subpath and operation
functions remain compatibility entry points; no existing API was deleted.
New application code should navigate the models rather than assembling operation
helpers. The factory's checked generic never asserts a class from erased types.

The model is a reviewed vertical slice, not complete legacy-format read/write
parity. Native snapshots test declared fields; they do not establish rendering,
all rich runs, media/objects or every unsupported record's semantics.
