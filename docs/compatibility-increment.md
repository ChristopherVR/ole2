# Compatibility coverage and evidence

These changes extend the published 0.13.0 baseline
(`c7bb2eeff47b9c631c5b0c61020bd9fcc4f699c2`) without changing the package
version. Existing native snapshots remain evidence only for their recorded
operations. New verification was run in Linux; it includes no new Windows
Office save/reopen validation.

| Format/version | Operation | New behavior | Verification | Remaining boundary |
| --- | --- | --- | --- | --- |
| CFB v3 rebuild | List, edit and rebuild directory metadata; repair | Exact unsigned FILETIME ticks survive instead of losing sub-millisecond precision; pre-1970 dates floor correctly | Zero, pre-epoch, sub-ms and maximum ticks; timestamp edits and invalid-range refusals; strict parser and packed consumer | Rebuild remains v3; this does not expand v4 writes or directory-name normalization |
| XLS BIFF8 | Replace existing RK/MULRK with a finite value outside RK encoding | Promote to exact IEEE NUMBER, splitting packed siblings and checking BOUNDSHEET/INDEX/DBCELL relocation | Native-generated RK/MULRK inputs; raw sibling/XF preservation; independent xlrd values/types/styles comparison; mixed retained handles and malformed pointer refusals | No native Excel acceptance of these new outputs; no formula execution or unchecked pointer relocation; explicit cell creation is bounded as described below |
| DOC Word 97 binary | Same-length paragraph edit and supported paragraph growth | Refuse multi-paragraph field overlap even on the same-length fast path; validate section CP bounds and table aliases before growth | Derivative of an existing historical Word fixture (origin unverified), byte-preserving failure checks, valid growth followed by same-length edit | Bookmark growth stays refused; table structure and arbitrary feature-table relocation remain unsupported |
| PPT 97–2003 | Read active slide text; replace fixed-length text | Bounded record/text/slide reads and refusal of missing, malformed or shared TextHeader ownership | Native-supported existing edit suites plus authored ownership and resource-limit cases; packed consumer type check | Text growth, absent-run creation and arbitrary ownership reconstruction remain unsupported |
| VSD v11 | Inspect explicit geometry | Read per-header section flags and stored MoveTo/LineTo/ArcTo operands through immutable model snapshots | Compressed/uncompressed authored records, owner/sibling isolation, malformed operands; existing native fixture hidden-section and rectangle values | ArcTo and multiple-section coverage is authored; physical order only, no list-order/formula/inheritance evaluation, guides or geometry writes |

The increment retains compatibility APIs and opaque streams. Refused edits
return the original bytes or throw the model's explicit unsupported-edit error
without committing state. Synthetic self-roundtrips and independent parser
agreement do not establish native parity.

Run `npm run typecheck`, `npm test`, `npm run build` and
`npm run test:package` from the repository. The packed check installs a tarball
in an isolated temporary consumer and compiles readonly/type-checked API use;
it does not publish anything.

## Verification

- TypeScript typecheck and build passed.
- Full Vitest suite: 81 files, 835 tests passed.
- Packed-package runtime and TypeScript consumer checks passed.
- `git diff --check` passed.
- xlrd 2.0.1 independently read promoted copies of `workbook-mulrk.xls`
  (25 slots, one sheet) and `workbook-features.xls` (4,388 slots, four
  sheets): exact PI target and unchanged remaining values, cell types and
  XF indices. `scripts/check-xls-numeric-promotion.py` reproduces this check
  when xlrd is already installed; it adds no package runtime dependency.
- Independent code review covered all five lanes and found no confirmed
  corruption or type-safety blocker.

## Additional bounded operations

Explicit cell creation and geometry inspection retain the existing unsupported
write boundaries.

| Format/version | Operation | Additional behavior | Evidence and limits |
| --- | --- | --- | --- |
| XLS BIFF8 | Explicit absent numeric-cell creation | `sheet.createNumericCell(row, col, value, xf)` creates a NUMBER after a verified numeric predecessor in an existing nonempty ROW/DBCELL block; ROW column extent and stream pointers update; returns a stable editable handle | Requires a selected existing cell XF and existing DIMENSIONS extent; refuses new rows, before-first placement, existing physical cells, merged/formula ranges, unsupported pointers and malformed blocks. Native-generated inputs and independent xlrd output checks; fresh native Excel save/reopen remains pending |
| VSD v11 | Stored Ellipse inspection | Geometry section rows include six readonly Ellipse operands: center, left and top points | Primary libvisio layout reference and compressed/uncompressed authored regressions, including truncation/nonfinite data. Native Ellipse acceptance is unverified; guide ownership and geometry writes stay refused |
| DOC Word 97 binary | Equal-length paragraph replacement | Refuses text pieces overlapping the FIB or referenced CHPX/PAPX pages; bounds formatting descriptor allocation | Malformed derivatives of the existing historical input and unchanged valid same-length edit behavior across four existing fixtures. No new bookmark support or native Word acceptance claim |

The existing DOC corpus has empty bookmark tables in all seven inspected
fixtures. Bookmark-aware growth remains refused until its structures and
native acceptance can be verified. VSD guide ownership likewise lacks a
verified corpus/model and remains refused.

### Independent creation evidence

xlrd 2.0.1 read created copies of the existing Excel-generated fixtures:
`workbook-mulrk.xls` at row 3/column 2/XF 62 (25 cell slots), and
`workbook-features.xls` separately at row 2/column 6 and row 3/column 3/XF 15
(4,388 slots in each copy). In each case the originally absent cell reads as
PI/NUMBER with the selected XF; all other values, cell types, XF indices,
sheet order and dimensions match the input. Reproduce with
`scripts/check-xls-numeric-promotion.py ORIGINAL EDITED ROW COL --created --xf XF`.
No application native acceptance is implied.

Independent review also verified the final malformed-cell-order regression
against the actual fixture records and passed all 23 creation tests. Creation
checks enforce 64 MiB input and 100,000-record budgets before building the
physical record model. Combined packed consumers exercise creation, subsequent
value edits, retained existing handles, duplicate-creation refusal and readonly
Ellipse types.

Fixture provenance follows `test/fixtures/NATIVE-VALIDATION.md`: historical
DOC/PPT origins remain unverified; repository-owned native-generated fixtures
are distinguished from those historical inputs.
No new external historical fixture was added.

## Bookmark admission boundary

Bookmark-dependent growth remains refused after independent layout and
architecture review. [The prerequisites report](doc-bookmark-prerequisites.md)
records the pairing, duplicate-order, sentinel, selection and allocation rules,
the seven-fixture audit, and the required native fixture/snapshot evidence.
Two new authored in-memory regressions on an owned Word fixture verify growth
refusal before/after a bookmark, clean and dirty model atomicity, stable handles
and byte-identical table preservation during an allowed same-length edit. This
does not admit bookmark relocation or establish native acceptance.

The captured runtime coverage includes exact CFB timestamps, checked XLS numeric
promotion and explicit bounded cell creation, DOC field and structural safeguards,
bounded PPT text/ownership checks, and stored VSD ArcTo/Ellipse/section inspection.
No package release is implied by integrating these source changes.
