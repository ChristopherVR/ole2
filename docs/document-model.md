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
| DOC | Existing plain paragraph text; guarded growth/shrink outside balanced main-story fields; existing exclusive direct bold/italic, font-size, none/single/double underline and logical paragraph-alignment operands | No style resolution, new formatting records, paragraph insertion/removal, table/object model writes, field-code/result editing or unsupported CP-table shifts; processing budgets apply |
| XLS | Existing numeric/string cells; physical BLANK/MULBLANK to number/plain string/boolean/error; NUMBER/RK/MULRK to boolean/error and existing BOOLERR replacement and conversion to finite number/plain string | Missing cells, merged followers, formulas, recalculation, unsupported type changes and unsafe relocation records/layouts refuse; selected rich string becomes plain while retaining XF and other aliases |
| PPT | Existing active fixed-length slide and notes-body text; supported small-anchor rectangle/text-box bounds | Inline shape text is linked by validated identity; outline refs may remain unresolved. Group/mirror/inherited/rotated/large-anchor edits refuse; no notes creation, field replacement or general rich-run reconstruction |
| VSD v11 | Existing same-length UTF-16 shape text; exclusive top-level literal transforms within an unchanged stored block allocation | No older-version model, formula/style/master evaluation, fields or structural creation; block relocation, shared/overlapping blocks and dependent transform edits refuse |
| CFB | Supported existing regular/mini stream resizing with path/hierarchy preservation | No new directory entries, external DIFAT expansion or variable-length v4 mini transitions; container support is separate from Office fidelity |

Native Microsoft Visio 16.0.20430.20140 accepted the owned native baseline and
its byte-exact no-op serialization. Its own save/reopen preserved all captured
fields. The released 0.9.0 VSD text writer returned a corrupt file after an
equal-length `Hello\n\n` to `World\n\n` edit: relocating the page block was
rejected by Visio even though the CFB container and parser self-roundtrip passed.
The 0.9.1 containment refused that edit. The 0.10 encoded-fit writer admitted
World, but a later Jello edit in another native fixture was rejected by Visio.
The writer now preserves exact decoded block length as well as original encoded
allocation, using bounded token variations without added decoded bytes or page
and ancestor relocation. Five production text cases open and survive native
save/reopen, with only captured target text changed. Insufficient capacity,
shared blocks and compression budgets still refuse without changing bytes,
dirty state or revision. This establishes the tested text case, not native
acceptance for every admitted write. The authored
synthetic fixture was also rejected by Visio; its libvisio results are separate
parser-mechanics evidence. Stored native text retains a terminal paragraph LF
that the COM `Text` property omits. Positive master/layer fidelity and rendering
remain unverified.

`paragraph.directAlignment` reports a direct logical value: `'start'`, `'center'`,
`'end'` or `'justify'`. Undefined means inherited or undecoded alignment. A setter
replaces an exclusive existing modern PAPX operand; matching legacy mirrors
admit only center/justify and update both operands. Legacy-only, conflicting,
shared, opaque or direction-dependent mirrored values refuse. Paragraph handles
remain stable across edits. Inspection bounds aggregate SPRM allocation to
262,144 records. Native Word center/justify edits and save/reopen preserve the
other captured paragraph and character fields; restoration is byte-exact.

`cell.value` can fill an existing physical BIFF8 BLANK or one cell in MULBLANK.
The selected XF and neighboring blank records remain intact. This does not
create a missing cell or clear an existing value implicitly. Native Excel
normal-load and save/reopen comparisons cover number, plain string, boolean and
error values while retaining captured formatting, anchors and other sheets.

`slide.notesStatus` distinguishes `'present'`, `'absent'` and `'unsupported'`;
`slide.notesDiagnostic` explains unsupported notes. For present notes,
`slide.notes.texts` exposes validated text atoms, including their `role`,
`encoding` and `editRefusal`. Only an eligible `'body'` atom accepts a same-length
UTF-16 replacement with supported controls/codepage. Rich-run modifications and
mirrors remain unsupported; existing run records stay unchanged.
Unknown fields remain readable where validated and preserved. Malformed notes
do not prevent a valid slide model from opening. Overlapping live objects or
save-history metadata refuse writes. Native PowerPoint comparisons cover ASCII
and Unicode notes; native saves may regenerate note shape IDs, so saved output
is compared against an unchanged native-save control.

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

`doc.paragraphs[index].runs[runIndex].directUnderline` edits an existing exclusive
direct operand using `'none'`, `'single'` or `'double'`. Inherited, missing, opaque
or aliased operands refuse. Reacquire the run after an edit; no formatting record
is inserted. See [direct underline](doc-underline.md).

Existing XLS boolean/error cells can become finite numbers or plain strings
through `cell.value`. The selected XF and unrelated records, SST aliases and
formula caches are retained. Missing cells, formula cells, clearing, malformed
BOOLERR discriminators and unsafe relocation records refuse atomically. Formula
caches are not recalculated; `recalculationRequired` remains explicit.

The root `parseOle2` keeps the complete `Ole2File` inspection shape (`entries` and
`getStream`) while returning the document union. `parseCompoundFile` names the
original raw reader. The original `ole2-parser-read` subpath and operation
functions remain compatibility entry points; no existing API was deleted.
New application code should navigate the models rather than assembling operation
helpers. The factory's checked generic never asserts a class from erased types.

The model is a reviewed vertical slice, not complete legacy-format read/write
parity. Native snapshots test declared fields; they do not establish rendering,
all rich runs, media/objects or every unsupported record's semantics.
