# Legacy PowerPoint

*PPT · POWERPOINT 97–2003*

`parsePpt` returns a `PptDocument` with active slides, `OfficeArt` shape identities, explicit anchors and validated inline text references. Opaque binary content remains preserved.

## Actual shapes and small-anchor edits

slides[index].shapes exposes IDs, kind, flags, names, coordinate spaces and copied geometry snapshots. Bounds use exact master units, eight per `PowerPoint` point.

- Unmirrored top-level small-anchor rectangles and text boxes support bounded position and extent setters. Candidate geometry is reparsed before commit.
- Shared, grouped, inherited, rotated/flipped, unknown-property and OOXML-mirrored geometry refuses editing.
- Group and child coordinates remain local. Large `ClientAnchor` values are raw with unknown geometry; outline-to-shape text references can remain unresolved.

## Existing direct character runs

`text.runs` exposes immutable UTF-16 spans and existing direct font-size exceptions. `runsStatus` distinguishes decoded, absent and unsupported formatting; `runsDiagnostic` explains unsupported records.

- `run.directFontSizePoints` replaces a uniquely owned existing two-byte size in eligible inline text boxes with an integer from 1 to 4000 points.
- Undefined means unresolved inheritance. Missing operands, unknown formatting, shared persists, save metadata, mirrors and unsupported shape contexts refuse; no run splitting or style insertion occurs.
- Native `PowerPoint` character-font comparisons and save/reopen match an unchanged save control except the declared sizes; notes IDs normalize in both controls.

## Document model

Assign slides[index].texts[index].text or an eligible slides[index].notes.texts body atom, then serialize() after a supported edit.

- Replacement text retains the UTF-16 length, original encoding and control/field marker positions.
- Active save history is resolved; stale saves, unknown records and surrounding bytes are retained.
- Mirrored text, physical atoms shared by multiple active slide positions, and atoms without a unique valid `TextHeaderAtom` owner are refused. Duplicate text atoms in one header group remain inspectable but cannot be edited.
- `notesStatus` distinguishes present, absent and unsupported notes; `notesDiagnostic` explains unsupported decoding. Notes fields, rich-run modifications, creation and live-object/save-history overlaps refuse writes; existing run records stay unchanged.
- Unsupported edits throw before changing dirty state, revision, model values or bytes.

## Explicit limits

The model exposes a bounded shape and text slice, with explicit resource budgets and target-specific refusal reasons.

- No arbitrary text growth or shape/layout mutation is implied.
- Encrypted PPT cannot produce a successful checked model.
- The separate `buildPptFile` exporter consumes a neutral WDeck model; source-to-WDeck conversion and rendering remain in the viewer.

## Compatibility APIs

`readPptSlideTexts`, `editPptSlideText` and record utilities remain available. `readPptSlideTexts(bytes, limits)` accepts positive integer `maxRecords`, `maxTextBytes` and `maxSlides` limits, defaulting to 100,000 records, 16 MiB of text payloads and 10,000 slides. Text bytes are counted across active references before each string is decoded; unknown child records consume the record budget, and persist entries are bounded by `maxRecords`. New existing-file edits should use the document model.

The resource and ambiguous-owner regressions use existing generated/native fixtures and structurally modified copies. They add no new native PowerPoint save/reopen evidence. Arbitrary text growth still requires relocating records and updating dependent ranges, and remains unsupported.

## Example

```ts
import { parsePpt } from '@christophervr/ole2';

const document = parsePpt(pptBytes);
document.slides[0].texts[0].text = 'Native title updated';
const saved = document.serialize();
```
