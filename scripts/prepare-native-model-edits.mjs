import { mkdirSync, readFileSync, writeFileSync } from 'node:fs';
import { resolve, join } from 'node:path';
import { pathToFileURL, fileURLToPath } from 'node:url';
import { createHash } from 'node:crypto';

// Unified typed-model path only: no standalone record-edit functions.
const output = resolve(process.argv[2] ?? '.native-validation');
const moduleDirectory = resolve(process.argv[3] ?? 'dist');
const fixtures = fileURLToPath(new URL('../test/fixtures/', import.meta.url));
mkdirSync(output, { recursive: true });
const { parsePpt, parseDoc, parseXls } = await import(pathToFileURL(join(moduleDirectory, 'ole2-document.js')).href);
const load = name => new Uint8Array(readFileSync(join(fixtures, name)));
const hash = bytes => createHash('sha256').update(bytes).digest('hex');
const cases = [];
function save(id, fixture, input, model, assertions) {
  if (!model.dirty || model.revision < 1) throw new Error(`${id}: model edit was not committed`);
  const bytes = model.serialize();
  if (hash(bytes) === hash(input)) throw new Error(`${id}: serialization unchanged`);
  const path = join(output, `model-${id}.${id}`);
  writeFileSync(path, bytes);
  cases.push({ kind: id, api: `parse${id[0].toUpperCase()}${id.slice(1)}`, source: join(fixtures, fixture), sourceSha256: hash(input), output: path, outputSha256: hash(bytes), dirty: model.dirty, revision: model.revision, recalculationRequired: id === 'xls' ? model.recalculationRequired : undefined, assertions });
}
const pptInput = load('ppt/native-text.ppt');
const ppt = parsePpt(pptInput);
ppt.slides[0].texts[0].text = 'Native title updated';
ppt.slides[1].texts[0].text = 'Unicode β updated';
save('ppt', 'ppt/native-text.ppt', pptInput, ppt, ['ASCII title Native title fixture → Native title updated', 'Greek title Unicode Ω fixture → Unicode β updated', 'All other captured slide texts, notes, shapes and dimensions unchanged']);

const docInput = load('doc/main-field.doc');
const doc = parseDoc(docInput);
doc.paragraphs[0].text = 'A longer model replacement before an untouched literal field.';
save('doc', 'doc/main-field.doc', docInput, doc, ['First plain paragraph grows to declared model text', 'QUOTE field type, code, result and paragraph-relative anchor unchanged', 'All other captured paragraph/font/count fields unchanged']);

const xlsInput = load('xls/workbook-mulrk.xls');
const xls = parseXls(xlsInput);
xls.sheets[0].cell(0, 1).value = 'Model conversion 日本語 😀';
save('xls', 'xls/workbook-mulrk.xls', xlsInput, xls, ['Numbers!B1 numeric20 becomes Model conversion 日本語 😀', 'Only target Value2/nonformula Formula display changes', 'All other captured cells, fonts, numeric formats, formulas, comments and merges unchanged']);

const manifest = { schemaVersion: 1, moduleDirectory, serialization: 'typed-model.serialize()', evidence: 'generation-only; native comparison required separately', cases };
writeFileSync(join(output, 'typed-model-generation.json'), JSON.stringify(manifest, null, 2));
console.log(JSON.stringify({ generated: cases.length, output, nativeValidated: false }));
