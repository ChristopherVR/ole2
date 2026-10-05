---
layout: home

hero:
  name: ole2
  text: Legacy Office files, as typed documents
  tagline: Parse DOC, XLS, PPT and VSD files into editable TypeScript models and serialize them without losing the bytes you did not touch. No framework or runtime dependencies.
  actions:
    - theme: brand
      text: Get started
      link: /getting-started
    - theme: alt
      text: API reference
      link: /api
    - theme: alt
      text: GitHub
      link: https://github.com/ChristopherVR/ole2

features:
  - title: Checked, typed parsing
    details: parseOle2 returns a discriminated union of CFB, DOC, XLS, PPT and VSD documents. Ask for a specific format with expect and wrong files throw instead of being guessed.
    link: /formats/models
    linkText: Document models
  - title: Edits that preserve the rest
    details: Supported setters validate the edit, change only the targeted bytes and keep unknown records and streams. Unsupported edits throw before anything changes.
    link: /document-model
    linkText: Serialization contract
  - title: Word, Excel, PowerPoint and Visio
    details: Paragraphs and runs, sheets and cells, slides and shapes, Visio pages and shapes. Publisher files can be inspected.
    link: /formats/word
    linkText: Format guides
  - title: MS-CFB containers
    details: Read and build compound files, resize nested streams in place and read or update SummaryInformation metadata.
    link: /formats/containers
    linkText: Containers
  - title: PPT export
    details: Build a binary .ppt from a framework-neutral deck model, including pictures, notes, embedded objects and optional RC4 encryption.
    link: /ppt-export
    linkText: PPT export
  - title: Clear support boundaries
    details: Every format has a documented supported subset. The limitations page is the single place that lists what is not covered.
    link: /limitations
    linkText: Limitations
---

## Example

```ts
import { parseOle2 } from '@christophervr/ole2';

const document = parseOle2(bytes); // CFB | DOC | XLS | PPT | VSD
if (document.kind === 'xls') {
	document.sheets[0]!.cell(1, 0).value = 'Updated';
}
const saved = document.serialize();
```

See [Getting started](./getting-started.md) for installation and the [format guides](./formats/models.md) for what each format supports.
