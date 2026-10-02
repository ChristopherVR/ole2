export const apiPages = [
	{
		slug: 'word', title: 'Legacy Word', kicker: 'DOC · WORD 97–2003',
		description: 'Read main-body paragraphs from legacy Word DOC files and apply narrowly guarded paragraph text edits.',
		intro: 'The Word helpers read plain text from the main document body of binary .doc files. The edit API replaces one existing paragraph while carrying over its paragraph and first-run formatting.',
		sections: [
			{ title: 'What works', body: 'Use `readOleDocParagraphs` to get main-body paragraph strings. `writeOleDocParagraphEdit` appends replacement text and updates the piece table rather than shifting bytes it does not understand.', bullets: ['Paragraph formatting and the first run’s character formatting are carried to the replacement.', 'The writer supports files whose only populated character-position table is the required section table.', 'A failed or unsupported edit returns the original byte array unchanged.'] },
			{ title: 'Known limits', body: 'This is not a general Word parser or renderer. It does not expose tables, headers, footers, footnotes, comments, fields, bookmarks, or rich run structure through this API.', bullets: ['Edits are refused when character-position tables for footnotes, annotations, fields, or bookmarks are populated.', 'Paragraph breaks in replacement text are converted to spaces.', 'The API operates on existing main-body paragraph indexes; it does not add or remove paragraphs.'] },
		],
		example: `import { readOleDocParagraphs, writeOleDocParagraphEdit } from '@christophervr/ole2';

const paragraphs = readOleDocParagraphs(docBytes);
if (paragraphs) {
  console.log(paragraphs[0]);
  const edited = writeOleDocParagraphEdit(docBytes, 0, 'Updated paragraph');
  if (edited === docBytes) console.log('Edit was unsupported or unsafe.');
}`,
	},
	{
		slug: 'excel', title: 'Legacy Excel', kicker: 'XLS · BIFF8',
		description: 'Read whole BIFF8 Excel workbooks, preview a bounded first worksheet and perform guarded numeric and string cell edits.',
		intro: 'The Excel API reads whole workbooks with `readXlsWorkbook` (every sheet, values, cached formula results and decoded formula text, styles, merges, views, comments and hyperlinks), reads a rectangular preview from the first worksheet, and edits selected cell records. It works with BIFF8 `.xls` streams and OLE-wrapped workbooks.',
		sections: [
			{ title: 'What works', body: '`readOleXlsGrid` returns a bounded grid (50 rows by 26 columns by default). Numeric edits update existing NUMBER or RK cells; string edits maintain the shared string table and can rebuild a single worksheet when the structure is recognized.', bullets: ['Custom row and column bounds can be passed to the grid reader.', 'String edits can replace supported cells or add a cell when the single-sheet layout is safe to resize.', 'Unsupported edits return the exact original input bytes.'] },
			{ title: 'Known limits', body: 'This is a data preview and cell writer, not an Excel calculation engine or workbook renderer.', bullets: ['Only the first worksheet is exposed by the grid API.', 'Formula values are not evaluated; shared strings split across CONTINUE records may not all resolve.', 'Resizing is limited to single-worksheet workbooks. Multi-sheet files permit only safe in-place edits.', 'Numeric writes target existing NUMBER or RK cells; string writes are bounded by understood SST and worksheet structures.'] },
		],
		example: `import {
  readOleXlsGrid,
  writeOleXlsNumericCellEdit,
  writeOleXlsStringCellEdit,
} from '@christophervr/ole2';

const grid = readOleXlsGrid(xlsBytes);
const numeric = writeOleXlsNumericCellEdit(xlsBytes, { row: 0, col: 0, value: 42 });
const text = writeOleXlsStringCellEdit(numeric, { row: 0, col: 1, value: 'Ready' });
if (text === xlsBytes) console.log('No supported edit was applied.');`,
	},
	{
		slug: 'powerpoint', title: 'Legacy PowerPoint', kicker: 'PPT · POWERPOINT 97–2003',
		description: 'Write a legacy PowerPoint PPT file from the neutral WDeck model and traverse binary record streams.',
		intro: 'PowerPoint support centers on binary output. `buildPptFile` serializes a neutral `WDeck` model into an OLE compound file; record helpers support lower-level binary inspection.',
		sections: [
			{ title: 'What works', body: 'The writer emits legacy presentation records from the model in `legacy-ppt-writer`. Its model includes slides, shapes, text runs, pictures, groups, media, notes, hyperlinks, embedded objects, and selected master styles.', bullets: ['The writer is asynchronous and returns `Uint8Array` bytes.', 'An optional password requests legacy RC4-encrypted output.', 'Record-stream helpers parse bounded record headers and iterate children.'] },
			{ title: 'Known limits', body: 'This package does not import or render complete PowerPoint presentations. It does not convert a viewer document into `WDeck`; that conversion belongs to the PowerPoint viewer integration.', bullets: ['Writing uses the documented neutral writer model, not arbitrary source-PPT round trips.', 'Record traversal is a binary utility, not a slide-layout engine.', 'Legacy PowerPoint output has format-specific constraints; it does not provide modern PPTX handling.', 'Several slide masters (`WDeck.masters` with `WSlide.masterIndex`) follow [MS-PPT] but are not yet verified by reopening in PowerPoint.', 'Pictures must be PNG, JPEG, EMF, WMF or DIB; convert GIF, TIFF and SVG first. Video has no record and is written as its poster picture; embedded sound must be WAV.', 'Password-protected output uses RC4 CryptoAPI only.'] },
		],
		example: `import { buildPptFile } from '@christophervr/ole2/legacy-ppt-writer';

const bytes = await buildPptFile({
  widthEmu: 9144000,
  heightEmu: 5143500,
  slides: [{ shapes: [] }],
  pictures: [],
});
// Save bytes as a legacy .ppt file.`,
	},
	{
		slug: 'visio', title: 'Legacy Visio', kicker: 'VSD · VISIO',
		description: 'Validate structural facts in legacy Visio containers without claiming drawing or page parsing.',
		intro: '`inspectLegacyVisio` checks a legacy VSD compound file’s VisioDocument signature, version, and TrailerStream pointer bounds. It reports structural facts only.',
		sections: [
			{ title: 'What works', body: 'The inspector validates the header signature, version-dependent pointer layout, trailer type, and trailer range against the declared document size. It returns stream names and trailer metadata when those checks pass.', bullets: ['Supports the validated V5 and V6+ pointer layouts.', 'The generic metadata API can read or update an existing SummaryInformation text property when its value fits the allocated slot.', 'Invalid or unsupported structures return `undefined` from the inspector.'] },
			{ title: 'Known limits', body: 'Visio shapes, pages, geometry, and drawing relationships are not decoded or editable. A valid TrailerStream pointer is not a parsed drawing.', bullets: ['No page preview or rendering API is included.', 'No general VSD save or structural edit API is included.', 'Metadata edits do not change drawing content and do not add missing properties.'] },
		],
		example: `import { inspectLegacyVisio, readLegacyOfficeMetadata } from '@christophervr/ole2';

const inspection = inspectLegacyVisio(vsdBytes);
if (inspection) {
  console.log(inspection.version, inspection.trailerLength);
  console.log(readLegacyOfficeMetadata(vsdBytes));
}`, 
	},
	{
		slug: 'publisher', title: 'Legacy Publisher', kicker: 'PUB · PUBLISHER',
		description: 'Inspect known legacy Publisher container signatures and versions; page layout is not decoded.',
		intro: '`inspectLegacyPublisher` validates known Publisher Contents signatures and the additional stream set required for Publisher 2002. This is format identification and container inspection, not publication-page support.',
		sections: [
			{ title: 'What works', body: 'The inspector recognizes validated Publisher 97/2000 and 2002 container signatures and returns stream names plus the streams used to validate the result.', bullets: ['Publisher 2002 requires the Contents, Escher/EscherStm, and Quill/QuillSub/CONTENTS streams.', 'The generic metadata API can edit an existing SummaryInformation text property when the replacement fits its allocated space.', 'Unknown or incomplete signatures do not produce a successful inspection.'] },
			{ title: 'Known limits', body: 'The package does not decode publication pages, text frames, images, or layout and does not provide a Publisher document writer.', bullets: ['Inspection results do not imply page content was parsed.', 'Metadata editing is independent of publication content.', 'No rendering, page preview, or structural publication edits are available.'] },
		],
		example: `import { inspectLegacyPublisher, writeLegacyOfficeMetadata } from '@christophervr/ole2';

const inspection = inspectLegacyPublisher(pubBytes);
if (inspection) {
  console.log(inspection.version, inspection.validatedStreams);
  const updated = writeLegacyOfficeMetadata(pubBytes, 'title', 'New title');
  if (updated === pubBytes) console.log('Metadata was unchanged or unsupported.');
}`, 
	},
	{
		slug: 'containers', title: 'OLE2 containers', kicker: 'MS-CFB · NAMED STREAMS',
		description: 'Read and build OLE2 compound-file containers and safely replace existing nested streams.',
		intro: 'The OLE2 layer exposes compound-file directory entries and stream bytes. These primitives are shared by the format-specific helpers elsewhere in the package.',
		sections: [
			{ title: 'What works', body: '`parseOle2` reads a compound file from an ArrayBuffer. `buildOle2` serializes a map of stream paths and bytes. Path-based stream replacement edits an existing stream while preserving unrelated bytes and nested storage layout.', bullets: ['The reader handles regular and mini streams, FAT/DIFAT allocation, and directory metadata.', 'Root-level SummaryInformation properties can be read, and existing text properties can be updated when they fit their allocated slots.', 'Metadata writes preserve the rest of the compound file and do not create absent properties.'] },
			{ title: 'Known limits', body: 'The container layer understands stream storage, not the internal application content of every file. Use the format guide for the actual Word, Excel, PowerPoint, Visio, or Publisher operation.', bullets: ['A stream name can identify a likely format but does not prove the file’s content is supported.', 'Replacement helpers are intentionally bounded to recognized paths and safe writes.', 'Metadata editing supports existing standard text properties only; it does not expand streams.'] },
		],
		example: `import { buildOle2, parseOle2 } from '@christophervr/ole2';

const bytes = buildOle2(new Map([
  ['Example', new Uint8Array([1, 2, 3])],
]));
const file = parseOle2(bytes);
console.log(file.getStream('Example'));`,
	},
];
