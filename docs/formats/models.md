# Document models

*CHECKED PARSING · TRANSACTIONAL EDITS*

The primary API is `parseOle2`: a checked discriminated union of `CfbDocument`, `DocDocument`, `XlsDocument`, `PptDocument` and `VsdDocument`. Concrete helpers return preserving editable classes.

## Runtime-checked types

The generic is a format key, never an erased document-class cast. `parseOle2<K>` requires expect: K and validates the bytes.

- Default kind narrows the union; `parseDoc`/`parseXls`/`parsePpt`/`parseVsd` return concrete classes.
- Wrong formats throw; ambiguous or unsupported automatic decoding returns a CFB view with diagnostics.
- expect: cfb explicitly requests container inspection for any valid compound file.
- The new vsd union variant requires exhaustive switch consumers to handle that kind.

## Mutation and serialization

Supported setters validate candidate edits before committing. `UnsupportedOle2EditError` leaves bytes, dirty, revision and model values unchanged.

- All classes extend `Ole2DocumentBase` and own copied input bytes.
- serialize returns a detached preserved-byte copy and does not reset dirty state.
- Capabilities describe supported operations; a specific target can still be refused.
- Opaque data is retained; complete legacy format fidelity and arbitrary edits are not claimed.

## Example

```ts
import { parseOle2, parseDoc, parseXls, parsePpt } from '@christophervr/ole2';

const document = parseOle2(bytes);
if (document.kind === 'xls') document.sheets[0].cell(1, 0).value = 'Updated';
const ppt = parseOle2<'ppt'>(pptBytes, { expect: 'ppt' });
ppt.slides[0].texts[0].text = 'Native title updated';
const saved = ppt.serialize();
```
