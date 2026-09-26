# Working agreements

- This package owns MS-CFB/OLE2 and legacy binary Office codecs. Modern DOCX belongs in docx-viewer/packages/core, not this repository.
- Preserve public binary APIs consumed by pptx-viewer and docx-viewer. Add fixture-based regression coverage for parser/writer changes.
- Keep runtime dependencies framework-free and source modules focused.
- Run typecheck, tests, build and packed-package smoke before release. Never publish test fixtures or stale output.
- Releases use matching v<version> tags, npm trusted publishing and provenance. Do not replace a published version.
