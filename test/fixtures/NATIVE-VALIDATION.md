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
`{"path":"/sheets/0/cells/INDEX/value","value":2.5}` objects using JSON Pointer.
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

Native snapshots capture semantic evidence: Word paragraphs/style/font flags,
headers/footers, main-story field codes/results, footnotes and table/shape counts;
Excel cell values/formulas/formats and sheet/name/shape metadata;
PowerPoint shape text/geometry, notes, and slide dimensions. They do not prove
rendering fidelity, picture payload identity, all character runs, external objects,
or every unsupported record. Pair them with container unknown-stream preservation
tests and format-specific parser tests. Parser self-roundtrips alone are not native
fidelity. Generated snapshots/edits stay local and are excluded from npm packs.
