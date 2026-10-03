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
element types, depth over 32, or configurable limits: 8,192 entries, 64 MiB per
stream and 128 MiB total. It reads bounded chunks rather than allocating declared
stream sizes. Compare against independently supplied generator manifests and
expected changes with the existing snapshot comparator. This is native container
evidence only; it says nothing about Office models or rendering. The v4 corpus
check covers 43 metadata entries and 41 hashed streams, including a stream beyond
the first 32 directory slots and a mini-stream. Dedicated checks exercised entry,
per-stream and total-byte rejection plus an invalid eight-byte input header.

## Typed document model vertical slices

After building the unified typed factory and adapters, run:

```powershell
node scripts/prepare-native-model-edits.mjs .native-validation dist
```

The driver uses only `parsePpt`, `parseDoc`, `parseXls`, supported model setters and
`serialize()`. Owned fixtures cover two fixed-length ASCII/Greek PPT title edits,
DOC paragraph growth outside an unchanged literal main-story field, and Unicode
string conversion of a middle MULRK cell with distinct neighboring XF entries.
It records exact source/output hashes and dirty/revision state; generation alone
is explicitly not native evidence. XLS also records `recalculationRequired`.
Use the native snapshot script to compare generated files against their original
fixture baselines with explicit declared text/value changes. Field locations
should be captured for the DOC case. These are captured semantic checks; neither
the driver nor successful serialization claims rendered or full-format fidelity.

## Optional character and cell-type probes

Use `-CaptureWordCharacterFonts` (bounded by `-MaxWordCharacters`, default 4096)
for selected small DOC fixture edits. This records resolved native appearance per
character; direct Word SPRM interpretation remains a separate codec assertion.
`-CapturePptCharacterFonts` has the same purpose for top-level PPT shape text,
with a total `-MaxPptCharacters` limit of 4096. Default snapshots are unchanged.
`-CaptureCellTypes` records native Excel Value2 type, displayed Text and a scalar
local ISERROR probe, distinguishing error HRESULTs from ordinary numbers.

`scripts/libvisio-snapshot.mjs` runs an explicitly supplied trusted `vsd2raw`
executable or portable extracted Ubuntu packages under WSL. It captures complete
drawing callbacks, input/executable hashes and tool version with input/output and
timeout bounds. Example (packages already extracted, no system installation):

```powershell
node scripts/libvisio-snapshot.mjs --wsl .native-validation/libvisio-portable/root .native-validation/owned.vsd .native-validation/owned-library.json Ubuntu
```

Library callbacks are independent parser evidence, not native Visio acceptance
or rendered fidelity. Executable provenance and binary fixture content rights
must be established separately; a source repository license alone does not
establish the rights to an arbitrary contributed binary. Unverified public test
samples stay local and untracked. The timeout/output bounds do not guarantee a
library memory limit or establish malformed-input safety.

## Native gates for the 0.10 parity increment

On 2026-10-03, fresh owned Word, Excel, PowerPoint and Visio instances (executable
version 16.0.20430.20140) accepted 13 typed-model cases and their native saved
copies. Inputs were repository-owned native-generated neutral fixtures; macros were disabled,
inputs opened read-only, and only generated copies were saved. These are captured
semantic assertions, not rendered or full-format fidelity.

The final combined build at `6a0f1714a5735d376939936f32ef175233c307ab` reproduced all 13 native-gated artifacts
byte-for-byte. The local evidence report `legacy-parity-evidence.json` has SHA256
`edb67f8f787a380b67ce1ce772121df2f1b6e04252abe29cda0b7c2899d1b465` and records
all source, output, native-save and snapshot hashes plus 69 compiled module hashes.
The report lives in the evidence owner's ignored `.native-validation` directory;
it is not a shipped fixture or an assertion that generation alone proves fidelity.

Frozen components: DOC `183b248a9770633e24d880f2351feab3f828487b`, XLS
`60577e0a41368d5cf7f77e4346690766a41b2886`, PPT
`2a92f4d180edd2b4a48f5570af0cfa7f34b50a7b`, and VSD
`1d40b76d95d0b39432b8b3a3f5856a5c1508bee5`.

Input fixture SHA256 identities:

| Fixture | SHA256 |
| --- | --- |
| `doc/paragraph-alignment.doc` | `a5ea2e6931374327fb8f5c4be50c241a9a742d9c5ede2143dd358ea1982e0ad1` |
| `xls/workbook-blanks.xls` | `69f3fca43378000aa4a41aee28ad2d04322d67796115b5bbf3ed0e7f5140a295` |
| `ppt/native-text.ppt` | `2bae7c1501262b79d2d9840ba968a5ed7b4ee12facde52082d3b114772f9b4e0` |
| `ppt/wide-notes.ppt` | `86905ce486c3fae0d4d63bc3f8187ca3d2ac87b46a599a7c6963d337b189ad42` |
| `vsd/native-visio16-v11.vsd` | `c6c97822e7bb2cc3e96da9d7d35fe74ac16a7f2c90d19ad4c80967d2896e0974` |

DOC changed only middle-paragraph alignment center to justify and restored center.
All captured text, per-character fonts, styles, stories, counts and other paragraph
formatting stayed exact, including native save/reopen. The justify artifact SHA256
is `d9fc75cbb499ef17dad2ef9623fce8fc664024f9ebd765d140e4179211f6557f`;
restoring center reproduced the input bytes. Matching physical and logical PAPX
alignment slots were updated together; this does not establish directional edits.

Eight XLS cases converted existing formatted BLANK/MULBLANK cells to numbers,
Unicode strings, booleans or errors on two sheets. Exact target values/types were
asserted while captured non-target cells, formulas, fonts, formats, comments,
merges and dimensions remained unchanged. Native saved/reopened captures matched.

| XLS artifact | SHA256 |
| --- | --- |
| first | `60327dde2cd5e1ebd91020ee3afa052865364de9da22f39f71de8ff50303fa80` |
| middle | `64867717f60fd0e9c5cffe2e591fa261e73a5f0a622db4d3ed25afe5c36ef2a6` |
| last | `cdb0349f74b9d27ba426f2984ca9fc19f465e1c1785020fbdbc6e908796e280d` |
| single-number | `c36694963364c901a046527f8b9268a91c36d4124f1be9c5bb115d3ba4d308d3` |
| single-string | `1a433e75b8c27f265cd507b124f604c41ad9af7a80997d43d51fe3af0125d7bd` |
| single-bool | `974a7632c81b4e73b984a0e9172a27888ad78b4104344dac50610104c87e8638` |
| single-error | `13b1b22afc8c7264c8d22aa5b29c2931b22cc64ce288ee868b7abcbbd56b0d8a` |
| later | `a77c5fb200b6631c88936f3682caca0b33528a0abdcab83f8f2725d8f76cd7ba` |

Excel baseline, eight edits and eight reopens were explicitly opened with
`CorruptLoad=0` (`xlNormalLoad`), which does not request object-model recovery.
The native `RepairMode` property was unavailable and is recorded as null; no
`RepairMode=false` claim is made.

PPT ASCII notes changed both `notes retained.` suffixes to `notes modified.`;
artifact SHA256 is `ea4de78e3e5671f38a8cbcd6c8845d76d2ec8d9cdf6f1d6f65b00b29f4983643`.
UTF16 notes changed `東京 Ω 😀` / `Wide notes.` to `大阪 β 😀` / `Wide model.`
while retaining the paragraph break and surrogate pair; artifact SHA256 is
`01abcbfe1bdf12afa02cbf9cd1e61a661d8d489b56422b2689d7f25e078c8b20`.
All other captured slide/notes text, fonts, geometry, IDs and counts were exact
before native save. Native saved edits matched unchanged save controls except
for declared text. PowerPoint remapped six ASCII notes-shape IDs identically in
the unchanged control; native-save ID preservation is not established.

The VSD production model changed stored `Hello\n\n` to `World\n\n`; COM exposes
one terminal LF, so the native assertion was `Hello\n` to `World\n`. Artifact
SHA256 is `cc257f441f2d30fb30ea9b0281fcf4570c092532518a4894a28fee4ae9bf3bdb`.
Fresh native open/save/reopen retained all other captured page/shape text, cell
formulas, transforms, geometry, IDs, counts and six styles. Independent Windows
IStorage inspection retained metadata, stream sizes and all other stream hashes.
This is one flat page with zero masters and layers, not positive complex-diagram
coverage. The prior native-rejected relocation output remains a negative gate.

Reproduce candidate edits using each component's tracked preparation driver,
then use `scripts/native-office-bounded.ps1` with the optional paragraph, blank-cell,
notes and character-font captures described above. Office ownership is proved by
one newly created application process, not by attaching to existing documents.
The Visio snapshot harness separately uses documented `OpenEx` flags 458 and its
owned-instance process proof. Compare native save/reopen with unchanged save
controls where the consumer normalizes IDs or metadata.

## Native compressed-text safety follow-up

A new owned literal-rectangle fixture exposed a native corruption case in the
0.10 encoder: `Hello\n\n` to `Jello\n\n` produced artifact SHA256
`d2f460372dc9bf34a8920f9e5eb806520dc8e7da0699aa592f027ff0c672201a`,
which Visio16 rejected at `OpenEx` with HRESULT `-2032464833`. The accepted
0.10 World case remains valid narrow evidence; it did not establish that added
decoded zero padding was safe for other text patterns.

The text-only fix at `6ad98f3c892ce68e4bb5b225de7435c1bd0333df` preserves the
exact decoded block length and uses bounded token variations to fit the original
encoded capacity. It refuses edits when it cannot find a safe fit. This introduces
no transform-writer coverage.

Run `node scripts/prepare-native-vsd-text-safety-edits.mjs dist/index.js` after
building. The driver pins the owned inputs, records six compiled module hashes,
seven output hashes and revision state, and checks atomic unsupported-length
refusal. Generation alone is not native validation.

Five production cases passed fresh macro-disabled native open, native save and
reopen in Visio16.0.20430.20140. Only the declared COM text changed; all other
captured cells/formulas/results, geometry, text, IDs, counts and six styles stayed
exact. Native immediately saved Document.Version is recorded separately from its
reopened value. The literal fixture SHA256 is
`14496febd65e0be62aa8fde493c0f07e812ac6181165b18899b2c1c8ea37bec7`.

| Production artifact | SHA256 |
| --- | --- |
| literal Jello | `7cbe62c046c62662dda6edb17dc56828b3eb575060ff9868952b083dce52c832` |
| literal World | `054508166099ee37892c4f0fec32078c90dd0ecdb2526b579c50d6e8b967ce31` |
| literal Again | `f734f5a168f679f1af8313bc4f59015c8edd9221b45714180b0a7429880a0761` |
| literal Hallo | `58ae283ade5758e3723dc26ca351b36b2da493202393d35fc06de4c41a59e759` |
| original three-shape World | `1d8d6d2c214490d4676adc758e4b1e8a0d383449d2cdb146e9bec47c829050bc` |

Hello no-op serialization is byte-identical to the literal source. A repeated
Jello/World/Again model edit produces the same Again artifact tested natively.
Both fixtures have zero masters and layers; these gates establish captured
semantics for the stated text cases, not rendered or full-format fidelity.

## Seeded native text matrix on 0.10.1

The released-equivalent build `92a2750c9031dd97448b00dae7fd9a2807be6570` passed
20 additional native Visio16.0.20430.20140 text cases on the two owned native VSD
fixtures. Seed `0xc0deface` pins varied ASCII entropy, Japanese/Greek/Cyrillic BMP,
combining marks, one and two supplementary surrogate pairs, the Unicode label,
repeated edits and edit/revert. Exact UTF16 text and control positions were
compared alongside every other captured native field. Five representative cases
also passed native save/reopen. This adds no transform-writing coverage.

`node scripts/prepare-native-vsd-text-matrix.mjs dist/index.js` reproduces inputs,
outputs, 69 module hashes and the seed. It also asserts 33 atomic refusals for
length/control/NUL/unpaired-surrogate changes, clean and already-edited state,
and a shape without explicit text. Two serialized files after clean refusals
opened natively with exact baseline captures. Generation alone is not native
validation. The actual run used 27 bounded fresh owned instances, macro-disabled
OpenEx flags 458 and EventsEnabled=false; all process proofs confirmed cleanup.
These are captured semantic gates on flat owned fixtures with zero masters and
layers, not rendered or full-format fidelity.

## DOC underline and XLS boolean/error scalar conversions

The owned Word 97-2003 `doc/underline-runs.doc` fixture, SHA256
`b788afb59be842da7baa0f65470ea298dd950851c7a0a411aa56c70046488154`,
passed native Word 16 single-to-double-to-none-to-single changes on the selected
nine characters, including native save/reopen. Only selected underline changed
among the captured character fonts, paragraph formats, text, styles, stories and
counts. Double and none replace one existing byte; single restores source bytes.
`scripts/prepare-native-doc-underline-edits.mjs` reproduces these artifacts.

The owned BIFF8 `xls/workbook-cell-types.xls` fixture, SHA256
`6cf446bec16da6acca52b48c18521f53fef609e0a2483dd8437a7b9097a89387`,
passed six native Excel 16 conversions of its three explicit boolean/error cells
to `37.125` or plain `Converted 漢字 😀` (the code points are recorded by the
reproduction driver). `scripts/prepare-native-xls-boolerr-edits.mjs` produces
all six native-tested artifacts and checks atomic refusals before and after
edits; generation alone does not launch Excel or establish native validation.
The selected value/type changed; all
other captured values, formulas, saved caches, fonts, styles, merges, counts,
comments and dimensions remained exact. Native save/reopen passed. Explicit
`xlNormalLoad=0` was used; the application did not expose RepairMode, so no
RepairMode=false assertion is made. The library does not evaluate formulas.

These bounded existing-record edits preserve unrelated records and refuse
ambiguous ownership or unsafe relocation atomically. They establish the captured
semantics of these owned fixtures, not pagination, rendering or full DOC/XLS
parity. The VSD text matrix and refusal gates remain unchanged; this increment
adds no native VSD transform support.

## Direct PPT font runs and VSD drawing-order reads

The owned `ppt/native-text.ppt` source
`2bae7c1501262b79d2d9840ba968a5ed7b4ee12facde52082d3b114772f9b4e0`
passed native PowerPoint 16 title-size changes from 28 to 32 points and Unicode
title-size changes from 24 to 30. Output SHA256 is
`505d12c3569cba74b5b66a8df93960d1d0c1afdfdc33a598e73d06fcea24cc45`;
only two existing CFB bytes differ. All selected character sizes changed as
declared; other captured fonts, styles, text, geometry, notes and counts stayed
exact. Native save/reopen matched an unchanged native-save control except the
declared sizes; both controls normalized the same six notes-shape IDs. The
tracked `prepare-native-ppt-font-runs.mjs` driver reproduces the artifact without
launching PowerPoint. Final model lookup caching preserves these output bytes.

The owned `vsd/native-hierarchy.vsd` fixture contains two native v11 pages with
IDs 0 and 4, sizes 8-by-11 and 11-by-8 inches, a group, a master, a custom style,
two layers, Unicode text and multiple geometry sections. Native save/reopen
preserved all captured fields. Page drawing orders `[2, 1, 5, 6]` and `[1, 2, 3]`
and group-five children `[3, 4]` match the new validated ShapeList metadata.
The previous physical-order root list `[1, 2, 5, 6]` did not match native order.
Existing flat shape IDs, parent/master references and local stored coordinates
remain unchanged; no-op serialization is byte-exact. Master/style evaluation and
hidden secondary geometry remain unresolved. This establishes declared read
semantics, not drawing rendering or native transform admission.
