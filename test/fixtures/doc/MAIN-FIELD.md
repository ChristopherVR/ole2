# Owned main-story field fixture

`main-field.doc` is original synthetic content generated in desktop Word16 by
`generate-main-field-fixture.ps1`. Three paragraphs contain a plain first target,
one benign QUOTE field in paragraph2, and a plain third target. Stored other-story
character counts are zero. Native Word's primary header/footer DOM ranges expose
empty paragraph marks, which do not imply populated binary header stories.

The literal field has code ` QUOTE "Fixture field" ` and result `Fixture field`.
It has no external reference, macro, embedded object or source document. Arial12
is a consumer font dependency; no fonts are embedded. Generated content is owned
by this repository under Apache-2.0. Metadata is stripped and neutral properties
are stamped. Office timestamps prevent byte-identical regeneration; update the
manifest hash after inspecting any regenerated revision.

This fixture targets variable-length editing outside main-story field ranges.
Native before/after snapshots must compare the literal field code/result and all
captured untouched paragraph/font/count fields. That is semantic evidence, not a
complete field evaluation engine or rendered-page fidelity claim.
