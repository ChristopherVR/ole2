# OLE2 containers

*MS-CFB · NAMED STREAMS*

`parseOle2` returns a checked document union by default. Explicit expect: cfb selects a `CfbDocument` container view for any valid compound file; this does not imply Office record semantics are decoded.

## Container model

Use stream(path).bytes to inspect or replace an existing stream by full storage path, then serialize().

- V3 regular/mini growth, shrink, emptying and transitions preserve hierarchy, unknown streams and directory metadata.
- V4 regular-stream resizing validates 4096-byte sectors and directory counts.
- New storage entries, external DIFAT expansion and variable-length v4 mini edits remain unsupported.

## Detection and diagnostics

Automatic parsing falls back to a CFB view with explicit diagnostics for ambiguous or unsupported Office models. Checked Office expectations throw instead.

- Container stream edits do not repair offsets inside Office records.
- Compatibility entries/`getStream` remain; returned inspection data is detached from owned bytes.
- `parseCompoundFile` names the original raw reader.

Directory entries expose exact `createdFileTime` and `modifiedFileTime` bigint
ticks alongside the millisecond `Date` views. `listCompoundFile` and
`repairCompoundFile` retain those ticks, including sub-millisecond precision.
Changing a listed `Date` writes the new date; to unset a timestamp, clear both
its `Date` and raw FILETIME fields. Invalid or out-of-range timestamps are refused.

## Example

```ts
import { parseOle2 } from '@christophervr/ole2';

const container = parseOle2(bytes, { expect: 'cfb' });
const stream = container.stream(['Custom', 'Data']);
stream.bytes = new Uint8Array([1, 2, 3]);
const saved = container.serialize();
```
