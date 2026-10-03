# Independent native consumer validation

`provenance.json` pins the existing binary fixtures by SHA-256. The XLS generator
is repository-owned synthetic content saved with Excel 16 as BIFF8. Historical
DOC/PPT fixture origin is unverified; this manifest does not invent attribution,
redistribution permissions, or a licensed public-corpus claim. New external
fixtures require source URL, upstream license and reproducible acquisition notes.

`doc/header-field.doc` is newly generated, repository-owned synthetic Word 16
content with a plain first paragraph, a PAGE field in its header, and a QUOTE
field in the second main paragraph. Regenerate with
`pwsh -STA -NoProfile -File test/fixtures/doc/generate-doc-fixtures.ps1`.
Its equal-length first-paragraph edit was independently checked in Word 16.

On Windows with desktop Office installed, first build the package, then run:

```powershell
node scripts/prepare-native-edits.mjs .native-validation
pwsh -STA -NoProfile -File scripts/native-office-snapshot.ps1 -InputPath test/fixtures/ole-word-97.doc -OutputPath .native-validation/doc-before.json
pwsh -STA -NoProfile -File scripts/native-office-snapshot.ps1 -InputPath .native-validation/edited.doc -OutputPath .native-validation/doc-after.json
node scripts/compare-native-snapshots.mjs .native-validation/doc-before.json .native-validation/doc-after.json .native-validation/doc-changes.json
```

Repeat snapshots for `xls/workbook-features.xls` and `edited.xls`, or
`sample-deck.ppt` and `edited.ppt`. Changes are a JSON array of
`{"path":"/sheets/0/cells/INDEX/value","value":2.75}` objects using JSON Pointer.
Choose INDEX from the native baseline, and explicitly expect Excel's numeric
`formula` field to change with the edited numeric cell. PPT edits change a shape's
`text` field. Every captured field outside these exact changes must match.

The script opens explicit paths read-only, disables VBA before opening, avoids
recent-file registration and link updates in Word/Excel, supplies empty passwords,
and never activates embedded objects. Do not use it on unsolicited or untrusted
downloads. Password-protected XLS is excluded. It closes its own documents and
application in `finally`; existing Office sessions are not enumerated or killed.
It caps Excel used ranges at 10,000 cells. COM setup/open failures are failures,
never successful fidelity evidence. Automated callers should enforce their own
process timeout and report blocked/hung consumers separately.

For small rich-string corpus cases, add `-CaptureRichText` to both Excel snapshot
commands. This captures each UTF-16 character's font name, size, bold, italic,
underline and color. It fails when any string exceeds `MaxRichTextCharacters`
(default4096), so callers cannot silently omit large unverified rich strings.
To inspect only explicitly selected cells in a mixed corpus, supply
`-RichTextCells 'Aliases!1,0;Aliases!1,1'` using sheet name and zero-based row,col
coordinates separated by semicolons. Every requested cell must exist and contain
text, or the snapshot fails. Character font evidence covers only those declared
cells; all other normal snapshot fields are still captured.
Ordinary snapshots retain their previous fields and avoid expensive per-character
COM calls. Excel's `Formula` getter can return null for long string cells; native
baseline10,000/12,000-character strings already exhibit this behavior. Assert the
full `Value2` text and explicitly record the consumer's null formula field.

For main-story field scenarios, add `-CaptureFieldLocations` to both Word snapshot
commands. Besides code/result/type, it records the field code's containing
paragraph index and relative CP offset within that paragraph. Absolute offsets
are not compared when preceding text grows or shrinks. This verifies that a field
remains anchored in the same untouched paragraph rather than merely retaining its
code/result somewhere in the document. The script never updates fields.

Native snapshots capture semantic evidence: Word paragraphs/style/font flags,
headers/footers, main-story field codes/results, footnotes and table/shape counts;
Excel cell values/formulas/formats and sheet/name/shape metadata;
PowerPoint shape text/geometry, notes, and slide dimensions. They do not prove
rendering fidelity, picture payload identity, all character runs, external objects,
or every unsupported record. Pair them with container unknown-stream preservation
tests and format-specific parser tests. Parser self-roundtrips alone are not native
fidelity. Generated snapshots/edits stay local and are excluded from npm packs.

## Independent Windows compound-file oracle

`scripts/native-cfb-snapshot.ps1` uses the Windows `ole32` structured-storage
implementation to open explicit compound files read-only (`STGM_READ` with
`STGM_SHARE_DENY_WRITE`). It never instantiates Office or activates embedded
objects. It enumerates nested storages, hashes complete stream bytes, and records
native class IDs, state bits, sizes, creation/modification times and element types.
Root filesystem timestamps are excluded because they are outside serialized CFB
directory fidelity. File allocation locations are deliberately absent from the
native semantic snapshot.

```powershell
pwsh -STA -NoProfile -File scripts/native-cfb-snapshot.ps1 -InputPath supplied.cfb -OutputPath .native-validation/cfb-native.json
```

The oracle fails explicitly on enumeration/read errors, truncation, unexpected
element types, depth over32, or configurable limits: 8,192 entries, 64MiB per
stream and 128MiB total. It reads bounded chunks rather than allocating declared
stream sizes. Compare against independently supplied generator manifests and
expected changes with the existing snapshot comparator. This is native container
evidence only; it says nothing about Office models or rendering. The v4 corpus
check covers43 metadata entries and41 hashed streams, including a stream beyond
the first32 directory slots and a mini-stream. Dedicated checks exercised entry,
per-stream and total-byte rejection plus an invalid eight-byte input header.
