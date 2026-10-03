# Reproducing the bounded PPT shape cases

Build the package, then run `node scripts/prepare-native-ppt-shape-edits.mjs`.
Optional arguments select an output directory and built module directory. The
driver does not launch Office. It emits three edited decks, the generated classic
baseline, explicit JSON Pointer changes, and a SHA-256/provenance manifest. It
fails if any byte hash differs from the independently observed native case.

| Case | Source | Observed before and after, in PowerPoint points |
| --- | --- | --- |
| Title position/width | Owned `test/fixtures/ppt/native-text.ppt` | Left 30 to 40; top 25 to 30; width 650 to 660; height 41.25 unchanged |
| Title height | Same owned native fixture | Height 41.25 to 51.25; other geometry unchanged |
| Classic rectangle | First-party synthetic neutral writer, `OwnedRectangle` | Bounds (10,20,100,50) to (20,25,120,60) |

The synthetic baseline is not native-produced. Its generator explicitly declares
the already-present primitive type and anchor in FSP flags. It removes no mirrors
or unknown properties. Both its baseline and edited output were independently
opened in Microsoft PowerPoint 16.0. All three comparisons matched the declared
geometry changes and preserved every other captured shape geometry, text,
character font, notes, slide ID and count. These are captured semantic checks;
they do not establish rendering or full format fidelity.

For a new independent check, use `scripts/native-office-snapshot.ps1` with each
manifest source and output path, then run
`node scripts/compare-native-snapshots.mjs BEFORE.json AFTER.json CHANGES.json`.
Use the native snapshot helper's rich-text option to include character formatting
when available. Explicit fixture inputs only; the helper disables macros and opens
documents read-only. The generation manifest deliberately does not report a new
native check as having run.

All successful cases use 8-byte small client anchors. Runtime bounds are exact
master units, eight per point. Large client anchors retain raw values with
undecoded geometry. Grouped, inherited, rotated/flipped, mirrored and ambiguous
geometry remains preserved and refused for editing. The mirrored rectangle in
the native fixture is explicitly refused; it is not the classic writable case.
