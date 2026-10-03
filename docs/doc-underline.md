# Existing direct DOC underline

`DocDocument.paragraphs[index].runs[index].directUnderline` reads and edits an existing direct CHPX `sprmCKul` operand. Supported values are `none`, `single`, and `double` (Kul bytes 0, 1, and 3). Absence, style inheritance, piece PRMs, duplicate operands, opaque context and other underline styles return `undefined` or refuse mutation. No formatting record is inserted and styles or underline colors are not resolved.

```ts
const paragraph = document.paragraphs[1];
paragraph.runs[0].directUnderline = 'double';
const bytes = document.serialize();
```

Writes reuse the exclusive full physical-run ownership checks. Shared CHPX blobs, partial paragraph views, overlapping logical pieces including other stories, PAPX aliases and embedded controls refuse. A supported edit changes only the existing one-byte operand; unrelated stream and formatting bytes remain unchanged. Invalid primitive values throw `UnsupportedOle2EditError` without coercion or document mutation. No-ops keep a clean document clean. Run handles expire after a successful edit; obtain a fresh handle from the paragraph for the next edit.

The official [MS-DOC Character Properties](https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-doc/7022285b-9621-42e9-ad4d-4e02c115ef18) defines sprmCKul, and [Kul](https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-doc/f6d2ad93-87cf-4e46-a992-a38ccf2ef295) defines its enumeration. This API reports direct stored exceptions rather than resolved appearance or full document fidelity.

`prepare-native-doc-underline-edits.mjs` produces reproducible edits of the repository-owned fixture after a build. It does not launch Office or claim native validation itself. Independent native Word validation must compare per-character underline, all other captured fonts, paragraph formats, text, styles, stories and document counts, including save/reopen.
