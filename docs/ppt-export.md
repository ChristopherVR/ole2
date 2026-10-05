# PPT export

`buildPptFile` creates a PowerPoint 97-2003 `.ppt` from a framework-neutral `WDeck` model: text, shapes, pictures, notes, embedded objects and optional RC4 encryption.

```js
import { buildPptFile } from '@christophervr/ole2/legacy-ppt-writer';

const pptBytes = await buildPptFile({
  widthEmu: 9144000,
  heightEmu: 5143500,
  slides: [{ shapes: [] }],
  pictures: [],
});
// Save pptBytes as a .ppt file. Pass { password: '...' } as a second argument
// only when legacy RC4 output is required.
```

Converting a PowerPoint viewer document into `WDeck` is the job of `pptx-viewer-core`, which consumes this shared binary implementation. Creating a new deck is separate from editing an existing one with `parsePpt`.

## Masters

`WDeck.masters` writes one `MainMaster` per entry and `WSlide.masterIndex` points each slide at its own. A deck without `masters` writes one master from `masterStyles` and `master`.

## Pictures

BLIPs are `png`, `jpg`, `emf`, `wmf` and `dib`, matching PowerPoint's own 97-2003 save. Convert other sources first.

## Constraints

Writer constraints (picture formats, media, encryption scheme, modern-only content, multi-master verification) are listed under [Limitations](./limitations.md#ppt-export).
