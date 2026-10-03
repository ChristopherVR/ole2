# Legacy format coverage and validation

This package owns binary legacy Office codecs. Modern OOXML and viewer rendering
remain in their consumer repositories. Container correctness, format parsing,
format writing, existing-file preservation and native application behavior are
separate validation layers.

| Layer/format | Read coverage | Write/edit coverage | Remaining parity gaps |
| --- | --- | --- | --- |
| MS-CFB v3/v4 | Bounded directory, FAT/DIFAT, mini/regular stream reads | Flat v3 creation; exact nested stream replacement; preserving v3/v4 regular-stream growth/shrinkage with header-only DIFAT | New nested storage creation, mini-stream resizing/transitions, high 64-bit stream sizes and DIFAT expansion |
| MS-OLEPS properties | Standard SummaryInformation properties | Existing supported text slots only | Creating/expanding properties, complete property-set models |
| DOC 97–2003 | Piece-table main-body text and paragraph indexes | Guarded paragraph replacement; see API outcome reasons and fixture tests | Complete tables, runs, fields, headers/footnotes, drawings and arbitrary CP-dependent structural edits; fresh-document writer |
| XLS BIFF8 | All worksheets, values/cached formulas/token text, styles, merges, views, comments, links and names | Exact NUMBER/RK/MULRK edits; preserving existing LABELSST/RK string edits in worksheet tab order, including SST continuations and supported offset relocation | Formula creation/recalculation, unsupported pointer-bearing records and cell conversions, charts/drawings, general workbook creation; BIFF5 and earlier |
| PPT 97–2003 | Active-save slide ids and outline/inline text; bounded record/persist traversal | Neutral WDeck exporter; guarded fixed-length text replacement in existing files | Full source-to-model import, shape-to-outline reference resolution, arbitrary layout/rich text edits, synchronized OOXML mirrors |
| VSD V5/V6+ inspection layouts | Signature/version/trailer bounds and metadata | Existing metadata slots only | Drawing/page models and drawing content writer; current support is structural inspection |
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

## Next engineering priorities

1. Grow a reproducible, licensed corpus with independent consumers, including
   complex fields/runs, Excel SST continuations, reordered sheet tabs and
   incremental PPT saves. Keep unsupported and malformed-input outcomes explicit.
2. Extend preservation-based editors before general reconstruction: mini-stream
   allocation changes, richer DOC CP table updates, and further XLS
   cell conversions and verified pointer-bearing record handling.
3. Build complete format models and format writers incrementally; share codecs
   with viewer consumers without duplicating renderer or OOXML implementations.
   VSD page/shape support requires its own format lane rather than a metadata claim.

References: [MS-CFB](https://learn.microsoft.com/en-us/openspecs/windows_protocols/ms-cfb/53989ce4-7b05-4f8d-829b-d08d6148375b),
[MS-DOC](https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-doc/ccd7b486-7881-484c-a137-51170af7cc22),
[MS-XLS](https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-xls/cd03cb5f-ca02-4934-a391-bb674cb8aa06),
[MS-PPT](https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-ppt/6be79dde-33c1-4c1b-8ccc-4b2301c08662).
