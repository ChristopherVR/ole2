# DOC objects

`readDocObjects(bytes, limits?)` reads the non-text content of a Word 97-2003 file
([MS-DOC]): inline and floating pictures, shapes, fields, bookmarks and embedded OLE objects. Every
position is an absolute character position (CP) across all stories, so it composes with a text or
formatting reader. It returns `undefined` for input that is not a readable unencrypted `.doc`, throws
`DocResourceLimitError` when a table exceeds `DocObjectsLimits`, and lists damage inside one optional
structure in `warnings` instead of failing the whole read.

| Result | Source |
| --- | --- |
| `inlinePictures` | 0x01 characters with `sprmCFSpec` and `sprmCPicLocation` into the `Data` stream (PICF plus OfficeArt BLIP) |
| `floatingShapes` | `PlcSpaMom`/`PlcSpaHdr` (FSPA) joined by shape id with the OfficeArt drawing in `DggInfo` |
| `blips` | the drawing group BLIP store, including out-of-line BLIPs in the `WordDocument` stream |
| `fields` | `PlcfFld*` for every story, nested, with typed HYPERLINK, REF, PAGEREF, TOC, PAGE, NUMPAGES, DATE, TIME, SEQ, MERGEFIELD and INCLUDEPICTURE |
| `bookmarks` | `SttbfBkmk` with `PlcfBkf` and `PlcfBkl` |
| `oleObjects` | characters with `sprmCFOle2`, resolved to `ObjectPool/_<id>` with ProgID, class id and preview picture |

PNG, JPEG and TIFF BLIPs are returned as stored, DIB BLIPs are wrapped as a BMP file, and EMF, WMF and
PICT metafile BLIPs are inflated and returned as raw metafile bytes with their bounds.

Not decoded: OfficeArt group transforms, gradient/pattern fills, shadows and effects (the raw property
table is on `DocFloatingShape.properties`), text box content (use `textboxId` with the text box story),
form-field data and hyperlink binary data (`HFD`), and linked-object monikers.

The object readers share one context with the formatting and structure readers (`openStoryContext`), and
`readDocModel` places their results inside the block tree: see [DOC model](/doc-model).
