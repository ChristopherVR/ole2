import { mkdirSync, readFileSync, writeFileSync } from 'node:fs';
import { resolve, join } from 'node:path';
import { pathToFileURL } from 'node:url';

// Deliberately fixed repository synthetic fixtures; never discover user documents.
const output = resolve(process.argv[2] ?? '.native-validation');
const moduleDirectory = resolve(process.argv[3] ?? 'dist');
const fixtureDirectory = resolve('test/fixtures');
mkdirSync(output, { recursive: true });
const load = name => new Uint8Array(readFileSync(join(fixtureDirectory, name)));
const module = name => import(pathToFileURL(join(moduleDirectory, name)).href);
const { writeOleDocParagraphEdit } = await module('ole-document-doc-editor.js');
const { writeOleXlsNumericCellEdit } = await module('legacy-excel-biff8.js');
const doc = writeOleDocParagraphEdit(load('ole-word-97.doc'), 2, 'Native consumer verified paragraph.');
writeFileSync(join(output, 'edited.doc'), doc);
writeFileSync(join(output, 'doc-changes.json'), JSON.stringify([{ path: '/paragraphs/2/text', value: 'Native consumer verified paragraph.\r' }], null, 2));
const xls = writeOleXlsNumericCellEdit(load('xls/workbook-features.xls'), { row: 1, col: 1, value: 2.5 });
writeFileSync(join(output, 'edited.xls'), xls);
console.log(JSON.stringify({ output, docBytes: doc.length, xlsBytes: xls.length, evidence: 'edit-generation-only' }));
// XLS expected paths must be chosen from the independent baseline snapshot,
// since sparse-cell ordering is a consumer property, not a parser assumption.
try {
  const { editPptSlideText } = await module('legacy-ppt-text.js');
  const result = editPptSlideText(load('sample-deck.ppt'), { slideIndex: 0, textIndex: 0, expectedText: 'Project\rAtlas', text: 'Project\rOrion' });
  if (result.status !== 'edited') throw new Error(`PPT edit rejected: ${JSON.stringify(result)}`);
  writeFileSync(join(output, 'edited.ppt'), result.bytes);
} catch (error) {
  if (error.code !== 'ERR_MODULE_NOT_FOUND') throw error;
  console.log('PPT edit module unavailable; DOC/XLS outputs generated.');
}
