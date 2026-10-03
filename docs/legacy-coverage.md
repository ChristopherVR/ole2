# Legacy format coverage and validation

This package owns binary legacy Office codecs. Modern OOXML and viewer rendering
remain in their consumer repositories. Container correctness, format parsing,
format writing, existing-file preservation and native application behavior are
separate validation layers.

| Layer/format | Read coverage | Write/edit coverage | Remaining parity gaps |
| --- | --- | --- | --- |
| MS-CFB v3/v4 | Bounded directory, FAT/DIFAT, mini/regular stream reads | Flat v3 container creation with valid directory trees; exact nested stream replacement | Generic nested storage creation/resizing; full directory metadata preservation when rebuilding |
| MS-OLEPS properties | Standard SummaryInformation properties | Existing supported text slots only | Creating/expanding properties, complete property-set models |
| DOC 97–2003 | Piece-table main-body text and paragraph indexes | Guarded paragraph replacement; see API outcome reasons and fixture tests | Complete tables, runs, fields, headers/footnotes, drawings and arbitrary CP-dependent structural edits; fresh-document writer |
| XLS BIFF8 | All worksheets, values/cached formulas/token text, styles, merges, views, comments, links and names | Exact NUMBER/RK/MULRK numeric edits; guarded first-sheet SST/cell edits | Formula creation/recalculation, complex string resizing, charts/drawings, general workbook creation; BIFF5 and earlier |
| PPT 97–2003 | Active-save slide ids and outline/inline text; bounded record/persist traversal | Neutral WDeck exporter; guarded fixed-length text replacement in existing files | Full source-to-model import, shape-to-outline reference resolution, arbitrary layout/rich text edits, synchronized OOXML mirrors |
| VSD V5/V6+ inspection layouts | Signature/version/trailer bounds and metadata | Existing metadata slots only | Drawing/page models and drawing content writer; current support is structural inspection |
| PUB 97/2000/2002 inspection | Recognized signatures and required streams | Existing metadata slots only | Publication/page models and content writer |

PPT text edits preserve character offsets and reject text that has an OOXML
`metroBlob` mirror. Unknown records and streams are retained by fixed-length
replacement. Rebuilding a format stream or a flat container has narrower
preservation guarantees and must refuse layouts it cannot safely update.

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

These snapshots cover the fields listed in
`test/fixtures/NATIVE-VALIDATION.md`; they do not establish complete layout,
rendering, run-level style, object or media fidelity. String-rebuild verification
is structural, not native fidelity evidence. Historical DOC/PPT provenance is
unverified in the manifest; the repository-owned Excel generator is reproducible.
No personal documents, embedded macros or objects are executed.

## Next engineering priorities

1. Grow a reproducible, licensed corpus with independent consumers, including
   complex fields/runs, Excel SST continuations, reordered sheet tabs and
   incremental PPT saves. Keep unsupported and malformed-input outcomes explicit.
2. Extend preservation-based editors before general reconstruction: nested CFB
   resizing, richer DOC CP table updates, and XLS string edits that retain every
   record and rebuild all affected offsets.
3. Build complete format models and format writers incrementally; share codecs
   with viewer consumers without duplicating renderer or OOXML implementations.
   VSD page/shape support requires its own format lane rather than a metadata claim.

References: [MS-CFB](https://learn.microsoft.com/en-us/openspecs/windows_protocols/ms-cfb/53989ce4-7b05-4f8d-829b-d08d6148375b),
[MS-DOC](https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-doc/ccd7b486-7881-484c-a137-51170af7cc22),
[MS-XLS](https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-xls/cd03cb5f-ca02-4934-a391-bb674cb8aa06),
[MS-PPT](https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-ppt/6be79dde-33c1-4c1b-8ccc-4b2301c08662).
