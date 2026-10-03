# Existing direct PPT character runs

`parsePpt(bytes).slides[index].texts[index].runs` returns immutable
`PptCharacterRun` handles with UTF-16 `start`, `end`, `text` and optional
`directFontSizePoints`. The implicit terminating paragraph mark is excluded from
the exposed span and preserved in the underlying style record. A missing size
means unresolved inheritance, not a default appearance.

```ts
const document = parsePpt(bytes);
const text = document.slides[0].shapes.find(shape => shape.name === 'FixtureTitle')!.texts[0];
if (text.runsStatus === 'decoded') text.runs[0].directFontSizePoints = 32;
const saved = document.serialize();
```

`runsStatus` is `decoded`, `absent` or `unsupported`; `runsDiagnostic` explains
unsupported formatting while valid slide text remains readable. Earlier versions
threw on `PptText.runs`. Consumers can now inspect these statuses and use the
readonly run list. The original three-argument `PptText` constructor remains
compatible, but a caller-created text atom has no document formatting owner.
Notes rich runs remain undecoded.

The setter changes only an existing two-byte font-size operand in an eligible
known inline text box. Values must be primitive integers from 1 to 4000 points.
Unknown styles, missing slots, terminal-mark-only runs, shared persist ownership,
save-history aliases, OOXML mirrors and unsupported shape contexts refuse
atomically. Outline text may have readable direct runs but remains unwritable.
No style record is inserted, split or cleared; effective inherited styling,
masters and other character properties are unresolved. Handles validate their
physical run identity and refresh through supported text/font changes.

The tracked `prepare-native-ppt-font-runs.mjs` driver reproduces the owned native
title and Unicode-title edits. Generation alone does not execute PowerPoint.
Native PowerPoint 16 comparisons cover all selected character sizes, other
captured fonts/styles/text, slide geometry, notes and counts. Save/reopen is
compared against an unchanged native-save control because both regenerate the
same six notes-shape IDs. These checks do not establish rendering or complete
rich-text fidelity.
