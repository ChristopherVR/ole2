# DOC model

`readDocModel(bytes, limits?)` composes the three rich Word 97-2003 readers into one block-tree model for a
consumer that maps it to WordprocessingML. It returns `undefined` for input that is not a readable unencrypted
`.doc`, throws `DocResourceLimitError` when a budget in `limits` is exceeded (and `Error` for an invalid limit), and
lists recoverable damage in `warnings`. It is read-only.

```ts
import { readDocModel } from '@christophervr/ole2';

const model = readDocModel(bytes); // DocModel | undefined
for (const block of model!.body.blocks) {
	if (block.type === 'paragraph') for (const item of block.inlines) { /* text, picture, fieldBegin, ... */ }
	else for (const row of block.rows) for (const cell of row.cells) { /* cell.blocks is a DocBlock[] */ }
}
```

## Shape

| Field | Content |
| --- | --- |
| `body` | the main story as `blocks` |
| `sections` | each `DocSection` with `headers` and `footers` by variant (`default`, `even`, `first`): `{ index, inherited, applies }`, where `index` points into `headersFooters` at the entry that supplies the content |
| `headersFooters` | one entry per section slot, with its story blocks (empty when the section inherits) |
| `noteSeparators` | the footnote and endnote separator, continuation separator and continuation notice stories |
| `footnotes`, `endnotes`, `comments`, `textBoxes` | the stories, indexed by the ids used in the inline items |
| `fields`, `bookmarks`, `blips` | the typed field list, bookmark list and picture store the inline items point into |
| `styles`, `lists`, `fonts`, `properties`, `facingPages` | from `readDocFormatting` and `readDocStructure` |
| `warnings` | deduplicated warnings of all three readers |

A story is `{ story, cpStart, cpEnd, blocks }`. A block is a paragraph or a table.

- `DocParagraphBlock`: resolved `props` (style cascade plus direct formatting), `markProps`, `style` (istd), `list`
  (`ilfo`, level, the override, the definition and the effective level, when the paragraph is numbered), `inlines`, and
  `endsSection` (the index of the section whose break this paragraph's mark carries).
- `DocTableBlock`: `gridTwips`, `rows` of `cells` with `gridColumn`, `columnSpan`, `rowSpan`, `vMerge`, cell and row
  `props` (borders, shading, widths, height) and `blocks`. Nested tables are blocks inside a cell. A `continue` cell is
  covered by the cell above that restarts the merge.

## Inline items

Control characters never appear in text. Tabs stay as `\t`; 0x1E and 0x1F become U+2011 and U+00AD.

| Item | Source |
| --- | --- |
| `text` | formatted characters, with the run's resolved `props` |
| `break` | 0x0B line, 0x0C page (not at a section end), 0x0E column |
| `picture` | 0x01 with a picture location; `picture.blip` has the image bytes |
| `ole` | an embedded object, placed after its field separator; `object.preview` is the preview picture |
| `fieldBegin`, `fieldInstruction`, `fieldSeparate`, `fieldEnd` | complex-field marks sharing an `id` into `model.fields`; `fieldBegin.field` is the typed field (HYPERLINK `url`, REF `bookmark`, ...). Fields may span paragraphs, so the marks are flat; the instruction text is `fieldInstruction`, the result is ordinary items |
| `bookmarkStart`, `bookmarkEnd` | positioned at the bookmark's CPs; `id` indexes `model.bookmarks` |
| `noteReference`, `commentReference` | 0x02 and 0x05 in the main story, with the id of the note or comment |
| `noteMark`, `separator` | the 0x02/0x05 mark at the start of a note or comment story; 0x03/0x04 in separator stories |
| `shapeAnchor`, `textBoxAnchor` | 0x08 in the main or header story, with the floating shape; `textBox` indexes `model.textBoxes` |

Header, separator and text box stories end with one extra bare paragraph mark that is not content; the model drops it.
A section break (0x0C at a section end) ends its paragraph and is reported as `endsSection`, not as a break item.

## Positions

Every `cp*` field is an absolute document CP. A `DocStoryRange` also carries the story id and story-relative
`start`/`end`. The table, section, note and text box readers keep their story-relative fields (`DocTable.start`, ...)
and document it.

## What is shared

The three readers run over one `StoryContext` (`openStoryContext`): one text decode, one piece table, one CHPX/PAPX run
index, one sprm parser (`parseGrpprl`), one sprm budget and one warning list. `DocModelLimits` is the union of the
budgets (`maxCharacters`, `maxPieces`, `maxParagraphs`, `maxRuns`, `maxSprms`, `maxCells`, `maxEntries`, `maxRecords`,
`maxBlipBytes`, `maxNesting`, ...); `DocStructureLimits`, `DocObjectsLimits` and `DocFormattingLimits` are subsets.
`readDocFormatting` now returns the paragraphs of every story in `stories` (`paragraphs` is still the main story).

## Not covered

- Symbol characters (`sprmCSymbol`), form-field data, hyperlink binary data and comment ranges (annotation
  bookmarks) are not decoded; revision marks, field-result caching flags and document protection are not mapped.
- Table row and cell properties beyond width, merge, vertical alignment, borders and shading stay in `props.sprms`.
- Floating shape content other than text boxes (groups, effects, gradients) is available only through the shape's
  raw `properties`.
- Page-level settings beyond the section properties (line numbering, footnote numbering formats, document grid) are not read.
