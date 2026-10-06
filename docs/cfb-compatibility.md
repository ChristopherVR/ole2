# CFB compatibility

The container reader accepts the irregularities that real writers produce, as long as accepting them cannot change stream content. Each accepted irregularity is reported in `warnings`. Stream bytes are never truncated, zero-filled or invented: a stream whose allocation cannot supply every advertised byte still throws.

| Irregularity | Default | `{ strict: true }` | Warning code |
|---|---|---|---|
| Byte-order mark other than `0xFFFE` | accepted | rejected | `byte-order` |
| Trailing bytes, or a trimmed final sector holding only padding | accepted | rejected | `unaligned-length` |
| Trimmed final sector holding stream bytes or allocation metadata | rejected | rejected | — |
| Uninitialized v3 stream-size high bits ([MS-CFB] 2.6.3) | ignored | ignored | `size-high-bits` |
| Mini-FAT sector count in the header disagrees with its chain | chain wins | rejected | `mini-fat-count` |
| Chain terminator malformed after every advertised byte was read | accepted | rejected | `chain-terminator` |
| Cyclic or dangling directory sibling/child link | link ignored | rejected | `directory-link` |
| Stale directory slot that no storage reaches | kept, `path` undefined | rejected | `unreachable-entry` |
| Stale slot with an undecodable name | skipped | rejected | `invalid-unreachable-entry` |

Reachable entries with undecodable names, cyclic FAT chains, DIFAT damage and out-of-bounds sectors are always rejected.

Edits stay strict. `replaceCompoundFileStream` and `resizeCompoundFileStream` refuse every irregular container. To edit one, rebuild it first with `repairCompoundFile`.

## Lookup

```ts
import { parseCompoundFile } from '@christophervr/ole2';

const file = parseCompoundFile(bytes); // Uint8Array or ArrayBuffer
file.getStream('WordDocument');                     // root stream wins over an embedded one
file.getStreamByPath('ObjectPool/_1234/Workbook');  // case-insensitive path
file.getStreamByPath('/SummaryInformation');        // matches \u0005SummaryInformation
file.findEntry('ObjectPool')?.modified;             // storage metadata
for (const entry of file.entries) entry.path;       // ['ObjectPool', '_1234', 'Workbook']
file.getStreamById(entry.id);                       // unambiguous, even with duplicate leaf names
```

Each entry has an `id` (its directory slot). `childId`, `leftSiblingId` and `rightSiblingId` refer to that slot, not to the position in `entries`. Entries also carry `parentId`, `color`, `stateBits`, `clsid`, `created` and `modified`.

## Writing, editing and repair

`buildCompoundFile` writes a v3 container with nested storages, per-entry CLSIDs, state bits and timestamps, and balanced, name-ordered sibling trees. Parent storages are created implicitly.

```ts
import { buildCompoundFile, listCompoundFile, repairCompoundFile } from '@christophervr/ole2';

const { nodes, options } = listCompoundFile(bytes);
nodes.push({ path: ['Attachments', 'note.txt'], data: noteBytes });  // add
const kept = nodes.filter((n) => n.path[0] !== 'Obsolete');         // delete
kept[0].path = ['Renamed'];                                         // move or rename
const rebuilt = buildCompoundFile(kept, options);

const { bytes: canonical, warnings, dropped } = repairCompoundFile(damaged, { dropUnreadable: true });
```

`repairCompoundFile` keeps every reachable storage and stream with its names, CLSIDs, state bits and timestamps. It discards stale slots and writes contiguous chains with exact header counts. A stream it cannot read throws, unless `dropUnreadable` is set; then the stream is omitted and its path is listed in `dropped`.

`buildOle2` remains the flat writer whose layout was validated against PowerPoint, and the PPT writer still uses it.

## PowerPoint text recovery

`salvagePptText(bytes)` returns `{ mode, slides, diagnostics }`, where `slides` is an array of text blocks per slide. It tries three modes in order:

1. `active`: the exact `readPptSlideTexts` model.
2. `persist-scan`: the newest UserEditAtom found by scanning the stream. This recovers files whose Current User offset is missing or wrong.
3. `record-scan`: every top-level slide container in stream order. Superseded saves may appear in this mode.

Encrypted documents are refused. Salvaged text is for reading and indexing only; use `readPptSlideTexts` for edit targets.

## Out of scope

The SheetJS `cfb` package also reads and writes ZIP archives and MIME "single file web page" (MHT) containers. They are not compound files, so this package does not handle them.
