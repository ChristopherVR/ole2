# Legacy format coverage and validation

This package owns binary legacy Office codecs. Modern OOXML and viewer rendering
remain in their consumer repositories. Container correctness, format parsing,
format writing, existing-file preservation and native application behavior are
separate validation layers.

| Layer/format | Read coverage | Write/edit coverage | Remaining parity gaps |
| --- | --- | --- | --- |
| MS-CFB v3/v4 | Bounded directory, FAT/DIFAT, mini/regular stream reads; CfbDocument path handles | Flat v3 creation; preserving v3 regular/mini growth/shrink and transitions; v4 regular growth/shrink, header-only DIFAT | New storage creation, variable-length v4 mini edits, high 64-bit sizes, external DIFAT expansion |
| MS-OLEPS properties | Standard SummaryInformation properties | Existing supported text slots only | Creating/expanding properties, complete property-set models |
| DOC 97-2003 | DocDocument main-body paragraphs, CP-mapped CHPX runs, direct formatting, PAPX style indices and logical paragraph alignment | Transactional paragraph text setters; guarded growth/shrink outside balanced main-story fields with CP relocation; existing exclusive bold/italic, half-point font-size, none/single/double underline and paragraph-alignment operands | Style resolution and new formatting records; mirrored alignment admits only direction-independent center/justify; editable table/object models, edits inside fields or other CP features/stories, paragraph creation and fresh-document writer |
| XLS BIFF8 | XlsDocument tab-order sheets/cells; value type, read-only workbook/style/formula/merge snapshots | Transactional numeric/string values; existing physical BLANK/MULBLANK to number/plain string/boolean/error; NUMBER/RK/MULRK to BOOLERR and existing boolean/error replacement and conversion to finite number/plain string; continued SSTs and supported offset relocation | Missing-cell/formula creation or recalculation, merged followers, string-to-BOOLERR, unsafe pointer records, charts/drawings, fresh workbook writer; BIFF5 and earlier |
| PPT 97-2003 | Active slides, OfficeArt shape IDs/kinds/flags/names, explicit anchor geometry, validated inline text, direct character-run spans/font sizes and linked speaker-notes text; bounded persist traversal | Fixed-slot slide and validated notes-body text; existing exclusive inline text-box font-size operands; conservative unmirrored top-level small-anchor rectangle/text-box bounds; separate neutral WDeck exporter | Outline shape refs may be unresolved; groups stay local; large anchors are raw; inherited formatting and other rich-run writes, notes fields and structural changes, and mirrored/inherited/rotated edits unsupported |
| Binary VSD v11 | Typed pages/shapes, dimensions/scale, explicit transforms, UTF-16 text and move/line geometry; bounded compressed pointers/records | Same-length text with bounded compression into the original allocation; exclusive supported literal transforms; no page/table relocation, original unknown records and CFB streams retained | Earlier versions, insufficient allocation capacity, general style/master/ShapeSheet evaluation, rich text, complex geometry and general native Visio fidelity; V5/V6+ structural inspection remains separate |
| PUB 97/2000/2002 inspection | Recognized signatures and required streams | Existing metadata slots only | Publication/page models and content writer |

PPT text edits preserve character offsets and reject text that has an OOXML
`metroBlob` mirror. Unknown records and streams are retained by fixed-length
replacement. Rebuilding a format stream or a flat container has narrower
preservation guarantees and must refuse layouts it cannot safely update.

V4 regular-stream resizing preserves nested hierarchy, raw directory metadata,
4096-byte header padding and unrelated physical bytes; it validates the declared
directory-sector count. Variable-length v4 mini edits and mini/regular transitions
remain refused. See [the v4 preservation scope](cfb-v4-preservation.md) for tested
and refused scenarios. Synthetic CFB checks establish container correctness,
separately from native application comparisons.

## Independent evidence

The primary API is the checked `parseOle2` document union with concrete format
helpers. Unsupported/ambiguous Office models fall back to an explicit CFB view
with diagnostics in automatic mode; checked expectations reject instead. See
[document models](document-model.md) for mutation, serialization and compatibility
contracts. Container operations and semantic Office models remain distinct.

The regression suite combines supplied binary Office files, generated records,
allocation corruption, directory tree invariants, save-history overrides and
exact unrelated-byte preservation. The packed npm consumer verifies the public
barrel and subpaths independently of the source checkout.

Windows native snapshots are available through
`scripts/native-office-snapshot.ps1` and `scripts/compare-native-snapshots.mjs`.
Word, Excel and PowerPoint 16 were actually launched on Christopher-PC for this
increment. Captured native comparisons passed for DOC paragraph replacement,
DOC modifier preservation, the four-sheet XLS B2 numeric edit, and PPT
`Product Overview` → `Product Snapshot` while retaining captured slide geometry
and notes. Numeric formula caches are not recalculated by this package.

The next preservation increment independently passed native comparisons for
semantically unchanged DOC/XLS/PPT regular-stream relocation, nested DOC
paragraph growth, Unicode and 24,000-code-unit continued-SST strings, and
rich-string/shared-alias preservation on a later XLS sheet. A repository-owned
PowerPoint fixture covers ASCII and Unicode text edits with reproducible
generation and recorded provenance. Original SST entries and unrelated BIFF
records are retained; BOUNDSHEET/INDEX and ExtSST pointers are updated or the
edit is explicitly refused. Excel formula caches are retained, not recalculated.

These snapshots cover the fields listed in
`test/fixtures/NATIVE-VALIDATION.md`; they do not establish complete layout,
rendering, complete run-level style, object or media fidelity. The newer XLS
rich-string case captures individual character fonts; it does not establish
general rich-text editing fidelity. Historical DOC/PPT provenance is
unverified in the manifest; the repository-owned Excel generator is reproducible.
No personal documents, embedded macros or objects are executed.

The next bounded write increment uses repository-owned native fixtures for DOC
paragraph alignment, physical XLS blanks, PPT notes and VSD text. Native Word
center/justify and eight Excel blank-value cases pass normal open/save/reopen
comparisons. PowerPoint ASCII and Unicode notes pass captured comparisons;
native saves regenerate some note shape IDs in both edited and unchanged-save
controls. Visio accepts the production `Hello\n\n` to `World\n\n` text output
and its native save/reopen, with only the captured target text changed. The
0.9.0 corrupt output and historical 0.9.1 safe refusal remain documented in
[the native Visio gate](visio-com-native-validation.md). These cases establish
the declared fields, not rendering, pagination or complete format fidelity.

An additional native fixture exposed an admitted Jello text edit rejected by
Visio in 0.10.0. The 0.10.1 writer preserves exact decoded block length as well
as encoded allocation, without adding decoded padding. Five production text
cases pass native open/save/reopen; edits without a bounded exact fit refuse.
PPT slide/outline text and geometry writes also refuse overlapping live persist
objects and save-history metadata. No-op edits retain bytes and clean state.

## Next engineering priorities

1. Grow a reproducible, licensed corpus with independent consumers, including
   complex fields/runs, Excel SST continuations, reordered sheet tabs and
   incremental PPT saves. Keep unsupported and malformed-input outcomes explicit.
2. Extend preservation-based models before general reconstruction: v4 mini-stream
   allocation changes, additional DOC CP tables, and further XLS
   cell conversions and verified pointer-bearing record handling.
3. Build complete format models and format writers incrementally; share codecs
   with viewer consumers without duplicating renderer or OOXML implementations.
   Extend VSD writes only with independent native transform/text mutation
   isolation and save/reopen evidence; formula/style/master evaluation remains
   a separate gap.

References: [MS-CFB](https://learn.microsoft.com/en-us/openspecs/windows_protocols/ms-cfb/53989ce4-7b05-4f8d-829b-d08d6148375b),
[MS-DOC](https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-doc/ccd7b486-7881-484c-a137-51170af7cc22),
[MS-XLS](https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-xls/cd03cb5f-ca02-4934-a391-bb674cb8aa06),
[MS-PPT](https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-ppt/6be79dde-33c1-4c1b-8ccc-4b2301c08662).
