# Native Visio COM gate, 2026-10-03

Microsoft Visio **16.0.20430.20140, AMD64** accepted the repository-owned
`test/fixtures/vsd/native-visio16-v11.vsd`. The fixture SHA256 is
`c6c97822e7bb2cc3e96da9d7d35fe74ac16a7f2c90d19ad4c80967d2896e0974`.
It contains one 8×11-inch page, three shapes (rectangle, one-dimensional line,
Unicode label ending in U+03A9), six styles, zero masters, and zero layers.
Its generator starts with a blank document and sets neutral fixture metadata.

The published 0.9.0 codec produced a **corrupt file** when shape 1 stored text
changed from `Hello\n\n` to `World\n\n`. Visio rejected output SHA256
`17da1ec00dc7a62bf1d4e84b8d43bc7729935b6b12b4986094f0fa99b14e2677`
with HRESULT `-2032466840`. This failed native result remains part of the evidence.

The reviewed containment fix `6eff0b937048c85b852706edb3568fde8bd0329b`
refuses that operation with `unsafe-block-relocation`. Both its no-op setter
and refused-edit serialization retain the original bytes, text, revision 0,
and clean state. **Fresh Visio instances accepted both serialized files.**
Their full captured snapshots matched the baseline: page dimensions/scale,
IDs/names/text, cell formulas and results, geometry row types/cells, child and
connection counts, master/layer arrays, and style count. Native baseline
SaveAsEx/reopen also preserved those captures. Refusal is not successful writing.

The independent local report records **19 cases, eight accepted native controls,
11 native rejections, and 22 named passing assertions**. Three safe root/size
relocation controls preserved the whole native snapshot; page/page-table
relocation controls failed even without semantic changes. The old handcrafted
fixture and its edits were rejected as an unrecognized version
(`-2032466854`), despite earlier independent libvisio acceptance.

This is the historical 0.9.1 containment gate. Current bounded-compression text
writes admit the tested World edit; this refusal driver requires the explicit
trusted 0.9.1 build, rather than the current dist:

```powershell
node scripts/prepare-native-vsd-refusal.mjs C:/trusted-ole2-0.9.1/dist/index.js
pwsh -NoProfile -File scripts/native-visio-snapshot.ps1 -InputPath .native-validation/native-visio/native-noop.vsd -OutputPath .native-validation/native-visio/noop.json
pwsh -NoProfile -File scripts/native-visio-snapshot.ps1 -InputPath .native-validation/native-visio/native-refused.vsd -OutputPath .native-validation/native-visio/refused.json
```

The 0.9.1 release integrates that containment fix at
`7592548eeb05b7e1577531400cade63e2a2fad7e`. It requires each encoded leaf to fit
its original allocation exactly and does not relocate page or ancestor blocks.

The driver can take an explicit trusted built `index.js` path and output
directory as its first two arguments. Its manifest records exact module and
artifact hashes. Generate fresh original content with
`test/fixtures/vsd/generate-native-fixture.ps1`; Office timestamps may vary.
The harness uses a new `Visio.InvisibleApp`, events disabled, and
[OpenEx flags 458](https://learn.microsoft.com/en-us/office/vba/api/visio.documents.openex)
(read-only, hidden, no MRU, macros disabled, no workspace).
[EventsEnabled set to false](https://learn.microsoft.com/en-us/office/vba/api/visio.application.eventsenabled)
also prevents RUNADDON execution during formula evaluation. Each case has a
45-second bound; cleanup uses only the owned HWND-derived Windows PID,
start time, and executable proof. All 19 recorded owned processes exited.

Limits: the stored model includes a terminal paragraph LF omitted by COM Text.
The binary version is 11; COM Document.Version returns 983040 on reopen and
720896 immediately after the explicit legacy save setting. The line transform
writer refuses `missing-or-formula-transform`. Zero masters/layers provide no
positive coverage for those features. No visual/render fidelity, evaluated
styles/formulas/masters, arbitrary geometry editing, or full VSD parity is claimed.
