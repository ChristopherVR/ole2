# Legacy Visio

*VSD · VISIO*

`parseVsd` returns a `VsdDocument` for the supported binary version 11 drawing slice. This uses legacy binary records; Microsoft MS-VSDX describes a different XML format.

## Typed drawing model

pages and shapes expose real IDs, stored page dimensions/scale, explicit transforms, UTF-16 text and `MoveTo`/`LineTo` geometry.

- Version 11 only; unsupported versions fail checked parsing or remain diagnostic CFB inspection in automatic parsing.
- Group/parent/master identities and unsupported geometry remain explicit; styles, masters and formulas are not evaluated.
- Compressed pointer/record traversal has input, decoded-byte, depth and record budgets.

## Validated drawing order

`page.topLevelShapeIds` and `shape.childShapeIds` resolve owner-scoped `ShapeList` mappings, with `shapeOrderIssue` when order is unresolved.

- Existing shapes arrays remain flat and in physical record order for lookup compatibility; use the new IDs for drawing order and retain group context.
- Copied frozen metadata matches a native two-page group/master/style/layer fixture. Parent and local-coordinate reads are verified; inherited rendering is unresolved.
- Malformed mappings, owners and ranges remain explicit. Aggregate order elements are bounded to 100,000; this adds no native transform writer admission.

## Preserving edits

Supported setters replace edited leaves at their original offsets and stored allocation sizes, retaining pointer ancestors, unknown records and other CFB streams.

- Text retains UTF-16 length and control positions; fields, shared or overlapping allocations refuse.
- Bounded compression preserves exact decoded length with no added bytes, at most 8 MiB decoded leaf/allocation and bounded token-fit search; insufficient capacity or budgets refuse atomically.
- Transform writes require an exclusive top-level literal transform with understood unit tags and no parent/master dependency.
- Literal path points are not automatically scaled or recalculated when transform dimensions change.

## Evidence and limitations

Native Visio 16 accepts five production exact-length text cases and native save/reopen with only captured target text changed. A newly rejected 0.10 Jello edit exposed unsafe decoded padding; the older handcrafted fixture is also rejected despite libvisio acceptance.

- Earlier binary versions, general `ShapeSheet` evaluation, styles, rich text, embedded content and general drawing reconstruction remain unsupported.
- `inspectLegacyVisio` still validates the signature/version/`TrailerStream` pointer for compatible inspection layouts; inspection is separate from drawing decoding.
- Native transform, rendering and complete roundtrip parity remain unverified; formula-bearing transforms refuse.

## Example

```ts
import { parseVsd } from '@christophervr/ole2';

const drawing = parseVsd(vsdBytes);
const shape = drawing.pages[0].shapes.find(shape => shape.id === 1);
shape.text = 'World\n\n'; // Native fixture stores Hello plus two LF characters.
const saved = drawing.serialize();
```
