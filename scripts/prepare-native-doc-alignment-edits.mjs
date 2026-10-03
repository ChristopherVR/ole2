import { readFileSync, writeFileSync, mkdirSync } from 'node:fs';
import { createHash } from 'node:crypto';
import { resolve, join } from 'node:path';
import { parseDoc } from '../dist/index.js';

// Owned neutral input only. Generation does not launch Office or execute objects.
const output = resolve(process.argv[2] ?? '.native-validation/doc-alignment');
const source = resolve('test/fixtures/doc/paragraph-alignment.doc');
const hash = bytes => createHash('sha256').update(bytes).digest('hex');
const input = new Uint8Array(readFileSync(source));
const sourceHash = 'a5ea2e6931374327fb8f5c4be50c241a9a742d9c5ede2143dd358ea1982e0ad1';
if (hash(input) !== sourceHash) throw new Error('Owned alignment fixture changed; obtain fresh native evidence');
const doc = parseDoc(input), paragraph = doc.paragraphs[1];
if (paragraph.directAlignment !== 'center') throw new Error('Unexpected direct alignment');
mkdirSync(output, { recursive: true });
const cases = [];
for (const alignment of ['justify', 'center']) {
  paragraph.directAlignment = alignment;
  const bytes = doc.serialize(), path = join(output, `alignment-${alignment}.doc`);
  const changes = [...bytes.keys()].filter(i => bytes[i] !== input[i]);
  if (changes.length !== (alignment === 'justify' ? 2 : 0)) throw new Error('Unexpected unrelated byte change');
  writeFileSync(path, bytes);
  cases.push({ alignment, nativeAlignment: alignment === 'justify' ? 3 : 1, path, sha256: hash(bytes), changes });
}
writeFileSync(join(output, 'manifest.json'), JSON.stringify({ source, sourceHash, generatorSha256: hash(readFileSync(new URL(import.meta.url))), nativeValidatedByThisRun: false, fullFidelity: false, paragraphIndex: 1, cases }, null, 2));
console.log(JSON.stringify(cases));
