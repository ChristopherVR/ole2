# Legacy Publisher

*PUB · PUBLISHER*

`inspectLegacyPublisher` validates known Publisher Contents signatures and the additional stream set required for Publisher 2002. This is format identification and container inspection, not publication-page support.

## What works

The inspector recognizes validated Publisher 97/2000 and 2002 container signatures and returns stream names plus the streams used to validate the result.

- Publisher 2002 requires the Contents, Escher/`EscherStm`, and Quill/`QuillSub`/CONTENTS streams.
- The generic metadata API can edit an existing `SummaryInformation` text property when the replacement fits its allocated space.
- Unknown or incomplete signatures do not produce a successful inspection.

## Known limits

The package does not decode publication pages, text frames, images, or layout and does not provide a Publisher document writer.

- Inspection results do not imply page content was parsed.
- Metadata editing is independent of publication content.
- No rendering, page preview, or structural publication edits are available.

## Example

```ts
import { inspectLegacyPublisher, writeLegacyOfficeMetadata } from '@christophervr/ole2';

const inspection = inspectLegacyPublisher(pubBytes);
if (inspection) {
  console.log(inspection.version, inspection.validatedStreams);
  const updated = writeLegacyOfficeMetadata(pubBytes, 'title', 'New title');
  if (updated === pubBytes) console.log('Metadata was unchanged or unsupported.');
}
```
