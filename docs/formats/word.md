# Legacy Word

*DOC · WORD 97–2003*

`parseDoc` returns a `DocDocument` with stable main-body paragraph handles. Supported text setters preserve opaque binary records and serialize the edited compound file without reconstructing unsupported content.

## Actual character runs

`paragraph.runs` maps CLX pieces and CHPX formatting pages to character positions. `styleIndex` reads the PAPX paragraph style identifier.

- Direct bold/italic/font-size/underline values are exposed separately from inherited or undecoded styling. Unknown SPRMs remain opaque.
- `directUnderline` replaces an existing exclusive `sprmCKul` operand with none, single or double; other underline styles and inherited slots refuse. Native Word save/reopen checks retain all other captured formatting.
- Existing exclusive understood bold/italic operands can be set to absolute values. `directFontSizePoints` can replace an existing exclusive `sprmCHps` slot with 1..1638 points in exact half-point increments. Shared blobs, opaque semantics and missing slots refuse.
- Run handles expire after edits; reacquire `paragraph.runs`. Styles and piece PRMs are not resolved.

## Direct paragraph alignment

`paragraph.directAlignment` exposes logical start, center, end or justify, separately from inherited styles.

- An existing exclusive modern PAPX alignment operand can be replaced; matching legacy mirrors admit center/justify and update both slots.
- Shared, opaque, conflicting or direction-dependent mirrored formatting refuses. Aggregate SPRM allocation is bounded to 262,144 records.
- Native Word center/justify and save/reopen comparisons preserve other captured formatting; restoration is byte-exact.

## Document model

Read paragraph.text and assign supported replacement text, then call serialize(). Successful edits mark dirty and increment revision.

- Equal-length plain targets retain runs and character-position tables when the encoding fits.
- Physical text ranges overlapping the FIB or referenced CHPX/PAPX pages refuse before patching. Formatting-page descriptor scans are bounded to 65,536 per table; malformed-fixture regressions verify atomic refusal.
- Field-range guards also apply to equal-length edits: a plain paragraph inside a multi-paragraph field result is refused, including fields omitted from the field PLC.
- Guarded growth/shrink outside balanced main-story fields shifts field positions while preserving codes, results and flags.
- Growth validates section CP bounds/order and refuses section or field tables aliased by another FIB table before changing any bytes. Derived malformed-fixture tests cover these guards; they do not add native Word validation evidence.
- Rejected setters throw `UnsupportedOle2EditError` before changing bytes or model state.

## Explicit limits

The paragraph collection has fixed structure. Rich runs expose a bounded direct-formatting slice; tables, fields, styles and drawings are preserved without complete editable models.

- No paragraph insertion/removal or embedded paragraph breaks.
- Edits inside field code/result ranges and unsupported CP-dependent features are refused.
- Bookmark-dependent growth remains unsupported; equal-length edits outside guarded field ranges preserve existing bookmark positions.
- [Bookmark prerequisites and fixture audit](../doc-bookmark-prerequisites.md) explain the retained refusal and the primary layout constraints.
- Processing budgets and supported container layout limits apply; no fresh-document writer is implied.

## Compatibility APIs

`readOleDocParagraphs` and `tryWriteOleDocParagraphEdit` remain available. New code should navigate the document model.

## Example

```ts
import { parseDoc } from '@christophervr/ole2';

const document = parseDoc(docBytes);
console.log(document.paragraphs[0].text);
document.paragraphs[0].text = 'A longer plain paragraph outside fields.';
const saved = document.serialize();
```
