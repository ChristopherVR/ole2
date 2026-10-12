# OLE2 containers

*MS-CFB · NAMED STREAMS*

`parseOle2` returns a checked document union by default. Explicit expect: cfb selects a `CfbDocument` container view for any valid compound file; this does not imply Office record semantics are decoded.

## Container model

Use stream(path).bytes to inspect or replace an existing stream by full storage path, then serialize().

- V3 regular/mini growth, shrink, emptying and transitions preserve hierarchy, unknown streams and directory metadata.
- V4 regular-stream resizing validates 4096-byte sectors and directory counts.
- Structural creation, deletion, renaming and movement are available for canonical v3 and v4 containers through storage and stream classes. These operations retain the original version and rebuild allocation; they preserve reachable stream contents and directory metadata, but not physical sectors, padding or slack bytes.
- Irregular containers are refused for structural edits. Preserving stream resizing still refuses external DIFAT expansion and variable-length v4 mini edits.

### Storage hierarchy and structural edits

`container.root` and `container.storage(path)` return `CfbStorage` handles.
`storage.children` is a readonly array of `CfbStorage | CfbStream`, discriminated
by `kind`. Storage navigation and creation take one child name; document
navigation and creation take an absolute array path. Parents must exist.

```ts
const attachments = container.root.createStorage('Attachments');
const note = attachments.createStream('Note', new Uint8Array([1, 2, 3]));
attachments.rename('Documents');
note.move([], 'RootNote'); // Absolute destination storage path; [] is the root.
attachments.delete();    // Recursively removes the storage and its descendants.
const saved = container.serialize();
```

Document-level `createStorage(path)`, `createStream(path, bytes)`, `delete(path)`,
`rename(path, name)` and `move(sourcePath, destinationPath)` expose the same
operations. Document `move` takes a complete destination entry path; handle
`move` takes a destination storage path and optional new name. Names are
case-insensitive for lookup and collision checks; case-only renames are allowed.
Root deletion/movement/renaming, replacement collisions and moves into a
storage's own descendants are refused.

Handles follow an entry and its descendants through rename/move. Unaffected
handles remain usable after rebuilding. Deleted handles permanently refuse
reads and edits even if another entry is created at the old path. `path` is an
immutable snapshot; request it again to observe a move. `entry` returns detached
directory metadata. Unsupported operations leave bytes, revision and dirty
state unchanged; an exact no-op rename/move stays clean.

These operations do not relocate references inside Office streams. Moving or
deleting format-required streams can make an Office document unusable even
though its CFB container remains valid. Format detection diagnostics describe
the original parse; structural edits keep this explicit CFB view.

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

`buildCompoundFile(nodes, { majorVersion: 4 })` creates a v4 container with
4096-byte sectors, a complete directory-sector count and zero header padding.
The default remains v3. `listCompoundFile` includes `majorVersion: 4` in the
returned build options for a v4 source, so list/build roundtrips preserve its
version. `repairCompoundFile` retains its existing canonical v3 output contract.
Strict reading and structural edits reject incorrect v4 directory-sector counts.

## Example

```ts
import { parseOle2 } from '@christophervr/ole2';

const container = parseOle2(bytes, { expect: 'cfb' });
const stream = container.stream(['Custom', 'Data']);
stream.bytes = new Uint8Array([1, 2, 3]);
const saved = container.serialize();
```
