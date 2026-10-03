import { mkdtemp, readFile, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join, dirname } from 'node:path';
import { spawnSync } from 'node:child_process';
import assert from 'node:assert/strict';
const directory = await mkdtemp(join(tmpdir(), 'ole2-package-'));
const npm = process.platform === 'win32' ? process.execPath : 'npm';
const npmArgs =
	process.platform === 'win32'
		? [join(dirname(process.execPath), 'node_modules/npm/bin/npm-cli.js')]
		: [];
function run(command, args, cwd = process.cwd(), capture = false) {
	const result = spawnSync(command, args, {
		cwd,
		encoding: 'utf8',
		stdio: capture ? 'pipe' : 'inherit',
	});
	if (result.error) throw result.error;
	if (result.status !== 0)
		throw new Error(`${command} failed (${result.status}): ${result.stderr || ''}`);
	return result.stdout;
}
const packed = JSON.parse(
	run(
		npm,
		[...npmArgs, 'pack', '--ignore-scripts', '--json', '--pack-destination', directory],
		undefined,
		true,
	),
)[0];
assert(packed.files.some((file) => file.path === 'dist/index.js'));
assert(
	!packed.files.some((file) => file.path.includes('docx')),
	'Modern DOCX code must not ship in the legacy package',
);
assert(!packed.files.some((file) => /^(?:src|test|node_modules)\//.test(file.path)));
const tarball = join(directory, packed.filename);
await writeFile(
	join(directory, 'package.json'),
	JSON.stringify({ name: 'ole2-consumer-smoke', private: true, type: 'module' }),
);
run(
	npm,
	[...npmArgs, 'install', '--ignore-scripts', '--no-audit', '--no-fund', tarball],
	directory,
);
const installed = JSON.parse(
	await readFile(join(directory, 'node_modules/@christophervr/ole2/package.json'), 'utf8'),
);
assert.equal(installed.name, '@christophervr/ole2');
await writeFile(join(directory, 'fixture.xls'), await readFile(new URL('../test/fixtures/xls/workbook-features.xls', import.meta.url)));
await writeFile(join(directory, 'fixture-blanks.xls'), await readFile(new URL('../test/fixtures/xls/workbook-blanks.xls', import.meta.url)));
await writeFile(join(directory, 'fixture.doc'), await readFile(new URL('../test/fixtures/doc/main-field.doc', import.meta.url)));
await writeFile(join(directory, 'fixture-size.doc'), await readFile(new URL('../test/fixtures/doc/rich-size-runs.doc', import.meta.url)));
await writeFile(join(directory, 'fixture-alignment.doc'), await readFile(new URL('../test/fixtures/doc/paragraph-alignment.doc', import.meta.url)));
await writeFile(join(directory, 'fixture.ppt'), await readFile(new URL('../test/fixtures/ppt/native-text.ppt', import.meta.url)));
await writeFile(join(directory, 'fixture.vsd'), await readFile(new URL('../test/fixtures/vsd/owned-v11.vsd', import.meta.url)));
await writeFile(join(directory, 'fixture-native.vsd'), await readFile(new URL('../test/fixtures/vsd/native-visio16-v11.vsd', import.meta.url)));
await writeFile(join(directory, 'fixture-literal.vsd'), await readFile(new URL('../test/fixtures/vsd/native-literal-transform.vsd', import.meta.url)));
await writeFile(
	join(directory, 'verify.mjs'),
	`
import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import { createHash } from 'node:crypto';
import { parseDoc, parseXls, parsePpt, parseVsd, parseCompoundFile, Ole2DocumentError } from '@christophervr/ole2';
const sizedDoc = parseDoc(readFileSync(new URL('./fixture-size.doc', import.meta.url)));
const sizedRun = sizedDoc.paragraphs.flatMap(p => p.runs ?? []).find(r => r.directFontSizePoints === 18);
assert(sizedRun);
sizedRun.directFontSizePoints = 13.5;
assert(parseDoc(sizedDoc.serialize()).paragraphs.flatMap(p => p.runs ?? []).some(r => r.directFontSizePoints === 13.5));
assert.throws(() => { sizedRun.directFontSizePoints = 13.25; });
import { editXlsPreservedStringCell, readXlsWorkbook } from '@christophervr/ole2';
import { editXlsPreservedStringCell as subpathStringEditor } from '@christophervr/ole2/legacy-excel-preserved-string-cell';
import { buildOle2, parseOle2, readCompoundFileStream, replaceCompoundFileStream, resizeCompoundFileStream, readOleXlsGrid, inspectLegacyVisio, inspectLegacyPublisher, writeLegacyOfficeMetadata, editXlsNumericCell, editXlsStringCell, readPptSlideTexts, editPptSlideText, tryWriteOleDocParagraphEdit } from '@christophervr/ole2';
import { buildPptFile } from '@christophervr/ole2/legacy-ppt-writer';
import { readRecord } from '@christophervr/ole2/legacy-ppt-record-stream';
import { readPptSlideTexts as subpathTextReader } from '@christophervr/ole2/legacy-ppt-text';

const bytes = buildOle2(new Map([['Sample', new Uint8Array([1,2,3])]]));
assert.deepEqual([...parseOle2(bytes).getStream('Sample')], [1,2,3]);
const changed = replaceCompoundFileStream(new Uint8Array(bytes), ['Sample'], new Uint8Array([3,2,1]));
assert.deepEqual([...readCompoundFileStream(changed, ['Sample'])], [3,2,1]);
const source = new Uint8Array(buildOle2(new Map([['Target', new Uint8Array(5000)], ['Unknown', new Uint8Array([7,8,9])]])));
const resized = resizeCompoundFileStream(source, ['Target'], new Uint8Array(9000).fill(5));
assert.equal(resized.ok, true);
assert.equal(readCompoundFileStream(resized.bytes, ['Target']).length, 9000);
assert.deepEqual([...readCompoundFileStream(resized.bytes, ['Unknown'])], [7,8,9]);
const mini = resizeCompoundFileStream(source, ['Target'], new Uint8Array([1,2,3]));
assert.equal(mini.ok, true);
assert.deepEqual([...readCompoundFileStream(mini.bytes, ['Target'])], [1,2,3]);
assert.deepEqual([...readCompoundFileStream(mini.bytes, ['Unknown'])], [7,8,9]);
const miniGrowth = resizeCompoundFileStream(mini.bytes, ['Target'], new Uint8Array(4096).fill(4));
assert.equal(miniGrowth.ok, true);
assert.equal(readCompoundFileStream(miniGrowth.bytes, ['Target']).length, 4096);
const refused = resizeCompoundFileStream(source, ['Target'], new Uint8Array(8 * 1024 * 1024));
assert.equal(refused.ok, false);
assert.equal(refused.bytes, source);
for (const api of [readOleXlsGrid, inspectLegacyVisio, inspectLegacyPublisher, writeLegacyOfficeMetadata, readRecord, editXlsNumericCell, editXlsStringCell, editPptSlideText, tryWriteOleDocParagraphEdit]) assert.equal(typeof api, 'function');
assert.equal(subpathTextReader, readPptSlideTexts);
assert.equal(subpathStringEditor, editXlsPreservedStringCell);
const workbook = new Uint8Array(readFileSync(new URL('./fixture.xls', import.meta.url)));
const stringValue = 'Unicode ' + String.fromCodePoint(0x03a9, 0x65e5, 0x1f600);
const stringEdit = editXlsPreservedStringCell(workbook, {row: 1, col: 0, value: stringValue});
assert.equal(stringEdit.status, 'edited');
assert.equal(readXlsWorkbook(stringEdit.bytes).sheets[0].cells.find(cell => cell.row === 1 && cell.col === 0).value, stringValue);
const xlsModel = parseXls(workbook);
xlsModel.sheets[0].cell(1,0).value = stringValue;
assert.equal(parseXls(xlsModel.serialize()).sheets[0].cell(1,0).value, stringValue);
assert.equal(xlsModel.dirty, true);
const blankModel = parseXls(new Uint8Array(readFileSync(new URL('./fixture-blanks.xls', import.meta.url))));
assert.equal(blankModel.sheets[0].cell(0, 0).type, 'blank');
blankModel.sheets[0].cell(0, 0).value = 3.25;
blankModel.sheets[0].cell(0, 1).value = true;
const blankSaved = parseXls(blankModel.serialize());
assert.equal(blankSaved.sheets[0].cell(0, 0).value, 3.25);
assert.equal(blankSaved.sheets[0].cell(0, 1).value, true);
assert.equal(blankSaved.sheets[0].cell(0, 2).type, 'blank');
assert.equal(blankSaved.sheets[1].cell(0, 0).value, 'Later text anchor');
assert.throws(() => parsePpt(workbook), Ole2DocumentError);
const docModel = parseDoc(new Uint8Array(readFileSync(new URL('./fixture.doc', import.meta.url))));
docModel.paragraphs[0].text = 'Model paragraph edit.';
assert.equal(parseDoc(docModel.serialize()).paragraphs[0].text, 'Model paragraph edit.');
const alignedModel = parseDoc(new Uint8Array(readFileSync(new URL('./fixture-alignment.doc', import.meta.url))));
assert.equal(alignedModel.paragraphs[1].directAlignment, 'center');
alignedModel.paragraphs[1].directAlignment = 'justify';
assert.equal(parseDoc(alignedModel.serialize()).paragraphs[1].directAlignment, 'justify');
const pptModel = parsePpt(new Uint8Array(readFileSync(new URL('./fixture.ppt', import.meta.url))));
pptModel.slides[0].texts[0].text = 'Native title updated';
assert.equal(parsePpt(pptModel.serialize()).slides[0].texts[0].text, 'Native title updated');
assert.ok(parseCompoundFile(pptModel.serialize().buffer).getStream('PowerPoint Document'));
assert.equal(pptModel.slides[0].notesStatus, 'present');
const noteBodies = pptModel.slides[0].notes.texts.filter(t => t.role === 'body');
assert.equal(noteBodies.length, 1);
assert.equal(noteBodies[0].text, 'Synthetic notes retained.');
noteBodies[0].text = 'Synthetic notes modified.';
const notesSaved = parsePpt(pptModel.serialize());
assert.equal(notesSaved.slides[0].notes.texts.find(t => t.role === 'body').text, 'Synthetic notes modified.');
assert.equal(notesSaved.slides[0].texts[0].text, 'Native title updated');
const vsdModel = parseVsd(new Uint8Array(readFileSync(new URL('./fixture.vsd', import.meta.url))));
assert.equal(vsdModel.kind, 'vsd');
const shape = vsdModel.pages[0].shapes[0];
assert.equal(shape.text, 'Hello\\n');
shape.text = 'World\\n';
shape.transform = { ...shape.transform, pinX: 6, width: 5 };
const savedVsd = vsdModel.serialize();
assert.equal(parseOle2(savedVsd).kind, 'vsd');
assert.equal(parseVsd(savedVsd).pages[0].shapes[0].text, 'World\\n');
assert.equal(parseVsd(savedVsd).pages[0].shapes[0].transform.pinX, 6);
assert.deepEqual([...vsdModel.getStream('OpaqueUnknown')], [9, 4, 8, 3, 5]);
assert.throws(() => parseVsd(workbook), Ole2DocumentError);
const nativeBytes = new Uint8Array(readFileSync(new URL('./fixture-native.vsd', import.meta.url)));
const nativeVsd = parseVsd(nativeBytes);
const nativeShape = nativeVsd.pages[0].shapes.find(s => s.id === 1);
assert.equal(nativeShape.text, 'Hello\\n\\n');
nativeShape.text = nativeShape.text;
assert.deepEqual(nativeVsd.serialize(), nativeBytes);
assert.equal(nativeVsd.dirty, false);
assert.equal(nativeVsd.revision, 0);
nativeShape.text = 'World\\n\\n';
const nativeSaved = nativeVsd.serialize();
assert.equal(createHash('sha256').update(nativeSaved).digest('hex'), '1d8d6d2c214490d4676adc758e4b1e8a0d383449d2cdb146e9bec47c829050bc');
assert.equal(parseVsd(nativeSaved).pages[0].shapes.find(s => s.id === 1).text, 'World\\n\\n');
assert.equal(nativeSaved.length, nativeBytes.length);
const nativeBeforeStream = readCompoundFileStream(nativeBytes, ['VisioDocument']);
const nativeAfterStream = readCompoundFileStream(nativeSaved, ['VisioDocument']);
assert.equal(nativeAfterStream.length, nativeBeforeStream.length);
assert.deepEqual(nativeAfterStream.subarray(0, 54), nativeBeforeStream.subarray(0, 54));
assert.equal(nativeVsd.dirty, true);
assert.equal(nativeVsd.revision, 1);
const literalBytes = new Uint8Array(readFileSync(new URL('./fixture-literal.vsd', import.meta.url)));
const literalVsd = parseVsd(literalBytes), literalShape = literalVsd.pages[0].shapes.find(s => s.id === 1);
literalShape.text = 'Jello\\n\\n';
const literalSaved = literalVsd.serialize();
assert.equal(createHash('sha256').update(literalSaved).digest('hex'), '7cbe62c046c62662dda6edb17dc56828b3eb575060ff9868952b083dce52c832');
assert.equal(parseVsd(literalSaved).pages[0].shapes.find(s => s.id === 1).text, 'Jello\\n\\n');
assert.equal(literalSaved.length, literalBytes.length);
assert.equal(literalVsd.revision, 1);
assert.throws(() => { literalShape.text = 'Too long\\n\\n'; });
assert.deepEqual(literalVsd.serialize(), literalSaved);
assert.equal(literalVsd.revision, 1);
assert.throws(() => { nativeShape.text = 'Too long\\n\\n'; });
assert.deepEqual(nativeVsd.serialize(), nativeSaved);
assert.equal(nativeVsd.revision, 1);
const invalid = new Uint8Array([1, 2, 3]);
assert.equal(editXlsNumericCell(invalid, { row: 0, col: 0, value: 1 }).bytes, invalid);
const ppt = await buildPptFile({ widthEmu: 9144000, heightEmu: 5143500, slides: [{ shapes: [] }], pictures: [] });
assert.ok(parseOle2(ppt.buffer).getStream('PowerPoint Document')?.length > 0);
assert.equal(readPptSlideTexts(ppt).slides.length, 1);
console.log('Packed legacy OLE2 package works in an independent npm consumer.');
`,
);
run(process.execPath, [join(directory, 'verify.mjs')], directory);
// Compile an independent consumer constructing the original public interfaces.
// Newly parsed metadata must remain additive for existing TypeScript callers.
await writeFile(join(directory, 'verify-types.ts'), `
import { readDocFib, readFcLcbAt, type DocFib, type DocCfbUnwrap } from '@christophervr/ole2';
import { parseOle2, parseDoc, parseXls, parsePpt, parseVsd, type Ole2File, type PptDocument, type VsdDocument, type DocParagraph } from '@christophervr/ole2';
const input = new Uint8Array();
const model = parseOle2(input);
const compatible: Ole2File = model;
if (model.kind === 'doc') model.paragraphs[0]!.text = 'paragraph';
if (model.kind === 'xls') model.sheets[0]!.cell(0,0).value = 'cell';
if (model.kind === 'ppt') model.slides[0]!.texts[0]!.text = 'fixed';
if (model.kind === 'vsd') model.pages[0]!.shapes[0]!.text = 'fixed';
const checkedVsd: VsdDocument = parseOle2<'vsd'>(input, {expect: 'vsd'});
// @ts-expect-error Checked VSD parsing requires its runtime expectation.
parseOle2<'vsd'>(input);
// @ts-expect-error VSD expectations cannot be substituted for another generic.
parseOle2<'vsd'>(input, {expect: 'ppt'});
// @ts-expect-error VSD page structure is fixed.
parseVsd(input).pages.push({id: 1, shapes: []});
const checked: PptDocument = parseOle2<'ppt'>(input, {expect: 'ppt'});
const inferred: PptDocument = parseOle2(input, {expect: 'ppt'});
// @ts-expect-error A generic format requires a runtime expectation.
parseOle2<'ppt'>(input);
// @ts-expect-error The explicit generic and runtime expectation must agree.
parseOle2<'ppt'>(input, {expect: 'xls'});
// @ts-expect-error Document classes cannot be used as erased cast generics.
parseOle2<PptDocument>(input, {expect: 'ppt'});
// @ts-expect-error Paragraph insertion is not supported by this model.
parseDoc(input).paragraphs.push({index: 0, text: 'new'});
// @ts-expect-error Model kind is immutable.
parseXls(input).kind = 'doc';
void compatible; void checked; void inferred; void parsePpt;
void checkedVsd;
const sizeHandle = parseDoc(input).paragraphs[0]?.runs?.[0];
if (sizeHandle) sizeHandle.directFontSizePoints = 13.5;
parseDoc(input).paragraphs[0]!.directAlignment = 'justify';
const originalParagraph: DocParagraph = { index: 0, text: 'Caller-created paragraph' };
void originalParagraph;
// @ts-expect-error Alignment names describe logical direction, not physical left/right.
parseDoc(input).paragraphs[0]!.directAlignment = 'left';
const parsedNotes = parsePpt(input).slides[0]!.notes;
if (parsedNotes) parsedNotes.texts[0]!.text = 'same-length notes';
// @ts-expect-error Notes ownership cannot be replaced by a caller.
parsePpt(input).slides[0]!.notes = undefined;
const fib: DocFib = {
  flags1Offset: 10, flags1: 0, tableStreamName: '1Table', cbMacOffset: 64,
  cbMac: 4096, ccpTextOffset: 76, ccpText: 0, plcfbteChpx: { fc: 0, lcb: 0 },
  plcfbtePapx: { fc: 0, lcb: 0 }, sed: { fc: 0, lcb: 0 }, clx: { fc: 0, lcb: 0 },
  fibRgFcLcbOffset: 154,
};
const container: DocCfbUnwrap = { wordDocBytes: new Uint8Array(), tableStreamName: '1Table',
  tableBytes: new Uint8Array(), rewrap: bytes => bytes };
const parsedVersion: number = readDocFib(container.wordDocBytes).nFib;
readFcLcbAt(container.wordDocBytes, fib, 0);
void parsedVersion;
`);
run(process.execPath, [join(process.cwd(), 'node_modules/typescript/bin/tsc'),
  '--noEmit', '--strict', '--target', 'ES2022', '--module', 'NodeNext',
  '--moduleResolution', 'NodeNext', join(directory, 'verify-types.ts')], directory);
console.log(`Verified ${packed.filename}; isolated consumer: ${directory}`);
