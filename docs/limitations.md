# Limitations

This is the single place that lists what the package does not support. The README intentionally does not repeat it. This is a codec library, not a complete legacy Office renderer, and format detection alone does not imply that content is parsed or editable.

General rules:

- An unsupported edit throws `UnsupportedOle2EditError` from the document model, or returns the original bytes / a `rejected` or `unchanged` status from the lower-level functions. Check the result before reporting success.
- Capabilities describe an editing surface, not the eligibility of every target; a specific cell, shape or paragraph can still be refused.
- Processing is bounded by input, decoded-size, depth and record budgets. Input beyond a budget is refused rather than partially processed.
- Internal Office record offsets are not repaired by container-level stream edits.

For the per-format coverage table and the validation behind it, see [Format coverage](./legacy-coverage.md).

## CFB containers

- Supported resize layouts: v3 containers with 512-byte sectors and at most 109 header-listed FAT sectors, and v4 regular streams with 4096-byte sectors.
- Not supported: creating new storage entries, external DIFAT expansion, variable-length v4 mini-stream edits and v4 mini/regular transitions, and high 64-bit stream sizes.
- Allocations are appended, so shrinking a stream does not compact the file.
- Only existing SummaryInformation text properties can be edited, and only when the value fits its slot. Creating or expanding properties and full property-set models are not supported.

See [CFB v4 preservation](./cfb-v4-preservation.md) for the tested and refused scenarios.

## Word (DOC)

- Only main-body paragraphs and a bounded direct-formatting slice (bold, italic, font size, underline, paragraph alignment) are modeled.
- No style resolution and no new formatting records.
- No paragraph insertion or removal and no embedded paragraph breaks.
- Edits inside field code or result ranges and other CP-dependent structures are refused; tables, objects and other stories have no editable model.
- Mirrored alignment admits only direction-independent center and justify.
- There is no fresh-document writer.

## Excel (XLS)

- BIFF8 only; BIFF5 and earlier throw `unsupported-version`.
- Formulas are never evaluated or written. Cached results are retained and `recalculationRequired` signals that the consuming application must recalculate.
- No creation of missing cells, no edits to merged followers, and no string-to-BOOLERR conversion.
- Records and layouts with unsafe pointer relocation are refused atomically.
- A selected rich string becomes plain text while other shared aliases keep their formatting.
- Charts, pictures, drawings, conditional formatting, data validation, pivot tables and VBA are listed as unsupported rather than decoded. There is no fresh-workbook writer.

## PowerPoint (PPT)

- Text edits keep the UTF-16 length, original encoding and control/field marker positions; arbitrary growth is not supported.
- Slides with an OOXML `metroBlob` mirror, and physical text atoms shared by several active slide positions, are refused.
- Shape geometry edits cover only unmirrored top-level small-anchor rectangles and text boxes. Shared, grouped, inherited, rotated or flipped geometry is refused.
- Direct character runs are read-only apart from existing inline text-box font sizes; there is no style insertion, run splitting or inheritance resolution.
- Notes fields, rich-run changes and note creation are refused. Outline-to-shape text references can stay unresolved, and group/child coordinates remain local.
- Encrypted files cannot produce a model.
- Layout, masters, pictures and animations are not decoded into an editable presentation model.

### PPT export

- **Several slide masters** are written following `[MS-PPT]` but the consecutive ids for masters after the first have not been verified by reopening in PowerPoint. The first id is verified for single-master decks.
- **Picture formats** are `png`, `jpg`, `emf`, `wmf` and `dib`. GIF, TIFF and SVG sources must be converted by the caller first.
- **Media**: embedded sound is written as a playable WAV. The model has no video record, so video and non-WAV audio should be passed as their poster picture.
- **Encryption** uses RC4 CryptoAPI only; the pre-CryptoAPI Office 95 scheme is not written.
- **Modern-only content**: charts, SmartArt, ink and 3D models have no binary record of their own. Callers write a preview picture and can attach the element's OOXML as a `metroBlob` so PowerPoint 2007+ reopens it as the native object.

## Visio (VSD)

- Binary version 11 only. Earlier versions are not decoded.
- Pages, shapes, stored transforms, UTF-16 text and move/line geometry are read. Styles, masters, general ShapeSheet and formula evaluation, rich text, embedded content, complex geometry and general drawing reconstruction are not.
- Text edits must fit the original allocation and exact decoded length. Insufficient capacity or work budgets refuse.
- Transform edits need an exclusive top-level literal transform with understood unit tags and no parent or master dependency. Literal path points are not rescaled when transform dimensions change.
- Child coordinates remain local and inherited rendering is unresolved.
- Native transform, rendering and complete round-trip parity are unverified. See [Visio native validation](./visio-com-native-validation.md).

## Publisher (PUB)

- Inspection only: known container signatures, versions and required streams.
- Publication pages, text frames, images and layout are not decoded, and there is no writer. Metadata editing is independent of publication content.
