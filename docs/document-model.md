# Typed editable legacy documents

The primary API parses a compound file into an editable document class. This
follows the viewer's document/content navigation pattern while retaining opaque
binary bytes rather than pretending the legacy formats have a complete model.

```ts
import {parseOle2, parseDoc, parseXls, parsePpt, parseVsd} from '@christophervr/ole2';

const document = parseOle2(bytes); // CFB | DOC | XLS | PPT | VSD document classes
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

All five classes extend `Ole2DocumentBase`. DOC owns stable paragraph handles;
XLS owns stable tab/cell handles and exposes read-only workbook, style, formula
and merge snapshots; PPT owns active slides with stable text/shape identities.
VSD owns version 11 pages and shapes with stored values. CFB
exposes existing stream handles by full storage path: `cfb.stream(['Custom',
'Data']).bytes = replacement`. This operation edits the container and does not
repair Office record offsets. No new directory entries are created.

## Checked parsing

The generic is a format key (`'cfb' | 'doc' | 'xls' | 'ppt' | 'vsd'`), not a document-class
cast. The overload requires `{expect: K}` for `parseOle2<K>` and validates that
expectation at runtime. Both `parseOle2<'ppt'>(bytes)` and an explicit `'ppt'`
generic combined with `{expect: 'xls'}` fail TypeScript compilation.
`parseDoc`, `parseXls`, `parsePpt` and `parseVsd` return concrete checked classes.
Adding `vsd` widens the default union; exhaustive switch consumers must handle it.

Default parsing returns a union discriminated by the immutable `kind` getter.
Ambiguous root format streams or a malformed/encrypted/unsupported Office model
return a `CfbDocument` with explicit `diagnostics` and, when uniquely recognized,
`detectedFormat`. This is container inspection, not successful Office decoding.
Checked Office parsing throws `Ole2DocumentError` instead of returning a false
model. `{expect: 'cfb'}` explicitly selects container inspection for any valid
compound file. Invalid CFB allocation/header parsing still throws.

VSD version 11 has a bounded stored-value drawing model. Earlier binary versions
remain unsupported for drawing decoding, with structural inspection available.
PUB page models remain unimplemented.

## Mutation and serialization

`capabilities.read`, `.write` and `.limitations` describe the supported surface,
not guaranteed eligibility of a particular target. Setters invoke the bounded
format writers, validate the candidate model and commit only supported changes.
An unsupported mutation throws `UnsupportedOle2EditError` with its reason before
changing bytes, model values, `dirty` or `revision`. A no-op leaves them unchanged.
Success marks `dirty` and increments `revision`. Existing handles observe later
supported edits rather than retaining stale snapshots. DOC character-run handles
are revision snapshots and expire after a successful edit; reacquire `.runs`.

`serialize()` returns a new byte array retaining all opaque data supported by the
underlying preservation strategy. It does not reset dirty state, write a file or
reconstruct unsupported content. Input bytes, returned streams, entry CLSIDs and
serialization output cannot alias owned bytes. Inspection snapshots are detached;
only supported model setters edit the document. Protected base methods are
internal adapter hooks, not a supported external mutation API.

| Model | Supported mutations | Explicit limitations |
| --- | --- | --- |
| DOC | Existing plain paragraph text; guarded growth/shrink outside balanced main-story fields; existing exclusive direct bold/italic and font-size operands | No style resolution, new formatting records, paragraph insertion/removal, table/object model writes, field-code/result editing or unsupported CP-table shifts; processing budgets apply |
| XLS | Existing numeric/string cells; NUMBER/RK/MULRK to boolean/error and existing BOOLERR replacement | No cell creation, formulas or recalculation; unsupported type changes and unsafe relocation records/layouts refused; selected rich string becomes plain while retaining XF and other aliases |
| PPT | Existing active fixed-length text; supported small-anchor rectangle/text-box bounds | Inline shape text is linked by validated identity; outline refs may remain unresolved. Group/mirror/inherited/rotated/large-anchor edits refuse; no general run/notes reconstruction |
| VSD v11 | Existing same-length UTF-16 shape text; exclusive top-level literal transforms within an unchanged stored block allocation | No older-version model, formula/style/master evaluation, fields or structural creation; block relocation, shared/overlapping blocks and dependent transform edits refuse |
| CFB | Supported existing regular/mini stream resizing with path/hierarchy preservation | No new directory entries, external DIFAT expansion or variable-length v4 mini transitions; container support is separate from Office fidelity |

Native Microsoft Visio 16.0.20430.20140 accepted the owned native baseline and
its byte-exact no-op serialization. Its own save/reopen preserved all captured
fields. The released 0.9.0 VSD text writer returned a corrupt file after an
equal-length `Hello\n\n` to `World\n\n` edit: relocating the page block was
rejected by Visio even though the CFB container and parser self-roundtrip passed.
The writer now refuses edits whose encoded block cannot retain its original
allocation, without changing bytes, dirty state or revision. This containment
does not establish native acceptance for every admitted write. The authored
synthetic fixture was also rejected by Visio; its libvisio results are separate
parser-mechanics evidence. Stored native text retains a terminal paragraph LF
that the COM `Text` property omits. Positive master/layer fidelity and rendering
remain unverified.

For DOC, `doc.paragraphs[index].runs[runIndex].directFontSizePoints = 13.5`
replaces an existing exclusive direct `sprmCHps` operand. Values must be primitive
finite numbers from 1 to 1638 points in exact 0.5-point increments, matching the
[MS-DOC character properties specification](https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-doc/7022285b-9621-42e9-ad4d-4e02c115ef18).
Missing, invalid or duplicate slots, opaque formatting, piece PRMs, partial runs
and shared or aliased formatting are refused. The setter preserves every byte
outside that operand. It does not insert font records, resolve inherited styles
or set complex-script font size. An undefined getter means inherited or
undecoded formatting. Invalid values throw `UnsupportedOle2EditError` with
`reason: 'invalid-formatting'`; expired handles throw `reason: 'stale-run'`.
After a successful mutation, reacquire the run through the stable paragraph's
`.runs` getter before another edit. No-op setters retain the handle and revision.

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
