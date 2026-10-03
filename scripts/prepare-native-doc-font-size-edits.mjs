import { readFileSync, writeFileSync, mkdirSync } from 'node:fs';
import { createHash } from 'node:crypto';
import { resolve, join } from 'node:path';
import { parseDoc } from '../dist/index.js';

// Repository-owned Word fixture only. This generator does not launch Office.
// Usage after build: node scripts/prepare-native-doc-font-size-edits.mjs [SOURCE] [OUTPUT]
const source = resolve(process.argv[2] ?? 'test/fixtures/doc/rich-size-runs.doc');
const output = resolve(process.argv[3] ?? '.native-validation/doc-font-size');
const hash = bytes => createHash('sha256').update(bytes).digest('hex');
const bytes = new Uint8Array(readFileSync(source));
const sourceHash = '07acf328ddc3e6c72126df3402a5f184088d3bf56e46c0e54d9a4245401f1b21';
if (hash(bytes) !== sourceHash) throw new Error('Source differs from the owned native18-point fixture');
mkdirSync(output, { recursive: true });
const records = [];
function save(doc, name, points) {
  const run = doc.paragraphs[1].runs[0];
  if (run.text !== 'Bold text') throw new Error('Unexpected target run');
  run.directFontSizePoints = points;
  const saved = doc.serialize();
  if (parseDoc(saved).paragraphs[1].runs[0].directFontSizePoints !== points) throw new Error('Candidate size mismatch');
  const path = join(output, `${name}.doc`);
  writeFileSync(path, saved);
  records.push({ name, path, sha256: hash(saved), points, cpStart: run.cpStart, cpEnd: run.cpEnd });
  return saved;
}
for (const [name, points] of [['size13-5', 13.5], ['size130', 130], ['size-min1', 1], ['size-max1638', 1638]]) save(parseDoc(bytes), name, points);
const doc = parseDoc(bytes);
const twelve = save(doc, 'size12', 12);
save(doc, 'size12-to13-5', 13.5);
const restored = save(doc, 'size12-to13-5-to12', 12);
if (hash(twelve) !== hash(restored)) throw new Error('12-point restoration did not preserve bytes');
save(doc, 'size-restore18', 18);
if (hash(doc.serialize()) !== sourceHash) throw new Error('18-point restoration did not preserve bytes');
writeFileSync(join(output, 'manifest.json'), JSON.stringify({ source, sourceHash, generatorSha256: hash(readFileSync(new URL(import.meta.url))), nativeValidatedByThisRun: false, fullFidelity: false, cases: records }, null, 2));
console.log(JSON.stringify(records));
