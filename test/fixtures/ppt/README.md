# Owned synthetic binary PowerPoint corpus

`native-text.ppt` contains two simple slides authored exclusively by
`generate-ppt-fixtures.ps1` in desktop PowerPoint16. The first uses compressed
Unicode text; the second has UTF-16 Greek, Japanese, accented text and a surrogate
pair. Both have synthetic notes; a plain rectangle supplies geometry to preserve.
There are no macros, embedded objects, links, charts or personal source documents.
The generated content is repository-owned and covered by the repository's
Apache-2.0 license. Installed Office themes/fonts remain consumer dependencies.

Regenerate with:

```powershell
pwsh -STA -NoProfile -File test/fixtures/ppt/generate-ppt-fixtures.ps1
```

The generator removes document metadata before stamping neutral properties.
Office embeds version details/timestamps, so regeneration is semantic rather than
byte-identical; update the provenance manifest's SHA-256 after inspecting a new
revision. It never opens any existing source documents or activates objects.

The generator uses [Microsoft's SaveAs file-type enumeration](https://learn.microsoft.com/en-us/office/vba/api/powerpoint.ppsaveasfiletype)
(`ppSaveAsPresentation=1`) and
[RemoveDocumentInformation](https://learn.microsoft.com/en-us/office/vba/api/powerpoint.presentation.removedocumentinformation)
with [`ppRDIAll=99`](https://learn.microsoft.com/en-us/office/vba/api/powerpoint.ppremovedocinfotype).

Parser regression tests check both encodings and exact-size guarded edits. Native
PowerPoint snapshots are independent evidence for captured text, notes and shape
geometry only; no full rendering, theme, image or unsupported-record fidelity is
claimed. Historical `sample-deck.ppt`/`picture-fixture.ppt` remain unchanged and
their original provenance remains explicitly unverified.
