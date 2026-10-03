import { readFileSync, writeFileSync, mkdirSync } from 'node:fs';
import { createHash } from 'node:crypto';
import { resolve, join } from 'node:path';
import { parseDoc } from '../dist/index.js';

// Owned neutral input only. No Office launch, macros or object execution.
const source = resolve(process.argv[2] ?? 'test/fixtures/doc/underline-runs.doc');
const output = resolve(process.argv[3] ?? '.native-validation/doc-underline');
const hash = bytes => createHash('sha256').update(bytes).digest('hex');
const input = new Uint8Array(readFileSync(source));
const sourceHash = 'b788afb59be842da7baa0f65470ea298dd950851c7a0a411aa56c70046488154';
if (hash(input) !== sourceHash) throw new Error('Owned underline fixture changed; obtain fresh native evidence');
const doc = parseDoc(input), paragraph = doc.paragraphs[1];
if (paragraph.runs[0].text !== 'Bold text' || paragraph.runs[0].directUnderline !== 'single') throw new Error('Unexpected direct underline run');
mkdirSync(output, { recursive: true });
const cases = [];
for (const underline of ['double', 'none', 'single']) {
  const run = paragraph.runs[0];
  run.directUnderline = underline;
  const bytes = doc.serialize(), path = join(output, `underline-${underline}.doc`);
  const changes = [...bytes.keys()].filter(i => bytes[i] !== input[i]);
  if (changes.length !== (underline === 'single' ? 0 : 1)) throw new Error('Unexpected unrelated byte change');
  if (parseDoc(bytes).paragraphs[1].runs[0].directUnderline !== underline) throw new Error('Candidate underline mismatch');
  writeFileSync(path, bytes);
  cases.push({ underline, nativeUnderline: underline === 'double' ? 3 : underline === 'single' ? 1 : 0, path, sha256: hash(bytes), changes, cpStart: run.cpStart, cpEnd: run.cpEnd });
}
writeFileSync(join(output, 'manifest.json'), JSON.stringify({ source, sourceHash, generatorSha256: hash(readFileSync(new URL(import.meta.url))), nativeValidatedByThisRun: false, fullFidelity: false, paragraphIndex: 1, cases }, null, 2));
console.log(JSON.stringify(cases));
