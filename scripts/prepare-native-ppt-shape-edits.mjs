import { mkdirSync, readFileSync, writeFileSync } from 'node:fs';
import { createHash } from 'node:crypto';
import { resolve, join } from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';

// Explicit repository-owned inputs only; generation does not launch Office.
// Usage after build: node scripts/prepare-native-ppt-shape-edits.mjs [OUTPUT] [DIST]
const output = resolve(process.argv[2] ?? '.native-validation/ppt-shapes');
const dist = resolve(process.argv[3] ?? 'dist');
const loadModule = name => import(pathToFileURL(join(dist, `${name}.js`)).href);
const { PptDocument } = await loadModule('ppt-document');
const { buildPptFile } = await loadModule('legacy-ppt-writer');
const { readPptSlideShapes, pptShapeChildren } = await loadModule('legacy-ppt-shape-reader');
const { readRecordOrThrow } = await loadModule('legacy-ppt-record-stream');
const { readCompoundFileStream, replaceCompoundFileStream } = await loadModule('ole2-stream-edit');
const { OA } = await loadModule('legacy-ppt-record-types');
const hash = bytes => createHash('sha256').update(bytes).digest('hex');
function assertHash(bytes, expected, label) {
  if (hash(bytes) !== expected) throw new Error(`${label}: bytes differ from the retained PowerPoint 16.0 case; obtain new native evidence`);
}
mkdirSync(output, { recursive: true });
const cases = [];
function save(id, input, doc, source, sourceHash, outputHash, changes, provenance) {
  assertHash(input, sourceHash, `${id} source`);
  const bytes = doc.serialize();
  assertHash(bytes, outputHash, `${id} output`);
  if (!doc.dirty || doc.revision !== 1) throw new Error(`${id}: expected one typed-model transaction`);
  const path = join(output, `${id}.ppt`);
  writeFileSync(path, bytes);
  const changesPath = join(output, `${id}-changes.json`);
  writeFileSync(changesPath, JSON.stringify(changes, null, 2));
  cases.push({ id, source, sourceSha256: hash(input), output: path, outputSha256: hash(bytes), changes: changesPath, provenance,
    changedBytes: input.reduce((count, value, index) => count + Number(value !== bytes[index]), 0) });
}
const nativeSource = fileURLToPath(new URL('../test/fixtures/ppt/native-text.ppt', import.meta.url));
const input = new Uint8Array(readFileSync(nativeSource));
const nativeHash = '2bae7c1501262b79d2d9840ba968a5ed7b4ee12facde52082d3b114772f9b4e0';
assertHash(input, nativeHash, 'owned native-text fixture');
const titleDoc = new PptDocument(input);
const title = titleDoc.slides[0].shapes.find(shape => shape.name === 'FixtureTitle');
if (!title || title.anchorKind !== 'client-small') throw new Error('Expected owned small-anchor title');
title.bounds = { ...title.bounds, x: 320, y: 240, width: 5280 };
save('title-geometry', input, titleDoc, nativeSource, nativeHash,
  'f381bdb878d11f433b6ccc8b431e1186e9b20ea773ec995b542dedd2302804b2',
  [{ path: '/slides/0/shapes/0/left', value: 40 }, { path: '/slides/0/shapes/0/top', value: 30 }, { path: '/slides/0/shapes/0/width', value: 660 }],
  'Repository-owned PowerPoint-created fixture; source provenance remains in fixture manifest');
const heightDoc = new PptDocument(input);
heightDoc.slides[0].shapes[0].height += 80;
save('title-height', input, heightDoc, nativeSource, nativeHash,
  'ec17e2e52474771f414ac9398e853502deb560dad9fc1d1ad69b38c3331ee8e4',
  [{ path: '/slides/0/shapes/0/height', value: 51.25 }],
  'Same owned title fixture; distinct height-only edit exercises native text-fit behavior');

// A first-party synthetic classic shape, NOT a native-produced document. This
// neutral writer emits explicit type/anchor records. Declare those existing
// records in FSP flags for the fixture; runtime edits never modify shape flags.
// No mirror blobs or unknown properties are removed to make an edit writable.
const created = await buildPptFile({ widthEmu: 9144000, heightEmu: 5143500, pictures: [], slides: [{ shapes: [{
  kind: 'shape', spt: 1, isConnector: false, name: 'OwnedRectangle',
  anchor: { x: 127000, y: 254000, w: 1270000, h: 635000 }, fill: { kind: 'solid', rgb: '4477AA' },
  text: { textType: 4, paragraphs: [{ indentLevel: 0, runs: [{ text: 'Owned rectangle' }] }] },
}] }] });
const stream = readCompoundFileStream(created, ['PowerPoint Document']);
const view = new DataView(stream.buffer, stream.byteOffset, stream.byteLength);
const record = readPptSlideShapes(created)[0].shapes[0];
const fsp = pptShapeChildren(view, readRecordOrThrow(view, record.headerOffset)).find(child => child.recType === OA.FSP);
if (!fsp) throw new Error('Synthetic fixture missing FSP');
view.setUint32(fsp.dataOffset + 4, 0xa00, true);
const classic = replaceCompoundFileStream(created, ['PowerPoint Document'], stream);
const classicHash = 'd145a810575649c9876e9f78bd8e2bce4791de484db67e82635d8b0957df70fd';
assertHash(classic, classicHash, 'classic baseline');
const classicSource = join(output, 'classic-rectangle-before.ppt');
writeFileSync(classicSource, classic);
const classicDoc = new PptDocument(classic);
classicDoc.slides[0].shapes[0].bounds = { x: 160, y: 200, width: 960, height: 480 };
save('classic-rectangle-after', classic, classicDoc, classicSource, classicHash,
  '89bab0d688fe63a19fa9c856137f62959ab084ba33fe13b86cde2f4bd5092ed7',
  [{ path: '/slides/0/shapes/0/left', value: 20 }, { path: '/slides/0/shapes/0/top', value: 25 },
    { path: '/slides/0/shapes/0/width', value: 120 }, { path: '/slides/0/shapes/0/height', value: 60 }],
  'First-party synthetic code-generated classic rectangle; explicit type/anchor flags; native baseline independently accepted');
writeFileSync(join(output, 'ppt-shape-generation.json'), JSON.stringify({ schemaVersion: 1,
  serialization: 'PptDocument.serialize()', evidence: 'generation-only; run separate native snapshots and comparisons',
  nativeValidatedByThisRun: false, fullFidelity: false, cases }, null, 2));
console.log(JSON.stringify({ generated: cases.length, output, retainedHashesMatched: true, nativeValidatedByThisRun: false }));
