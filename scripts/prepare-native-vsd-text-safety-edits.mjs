import { readFileSync, writeFileSync, mkdirSync } from 'node:fs';
import { createHash } from 'node:crypto';
import { resolve, dirname, join } from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';
import assert from 'node:assert/strict';
// Candidate generation only: native read-only snapshots/save/reopen are separate.
const root = resolve(dirname(fileURLToPath(import.meta.url)), '..');
const modulePath = resolve(process.argv[2] ?? join(root, 'dist/index.js'));
const output = resolve(process.argv[3] ?? join(root, '.native-validation/vsd-text-safety'));
const { parseVsd } = await import(pathToFileURL(modulePath).href);
const sha = bytes => createHash('sha256').update(bytes).digest('hex');
const sources = [
  ['literal', 'native-literal-transform.vsd', '14496febd65e0be62aa8fde493c0f07e812ac6181165b18899b2c1c8ea37bec7'],
  ['original', 'native-visio16-v11.vsd', 'c6c97822e7bb2cc3e96da9d7d35fe74ac16a7f2c90d19ad4c80967d2896e0974'],
];
mkdirSync(output, { recursive: true });
const cases = [];
for (const [kind, name, hash] of sources) {
  const sourcePath = join(root, 'test/fixtures/vsd', name), source = readFileSync(sourcePath);
  assert.equal(sha(source), hash);
  for (const words of (kind === 'literal' ? [['Jello'], ['World'], ['Again'], ['Hallo'], ['Hello'], ['Jello','World','Again']] : [['World']])) {
    const document = parseVsd(source), shape = document.pages[0].shapes.find(s => s.id === 1);
    assert.equal(shape.text, 'Hello\n\n');
    for (const word of words) shape.text = word + '\n\n';
    const bytes = Buffer.from(document.serialize()), last = words.at(-1);
    assert.equal(parseVsd(bytes).pages[0].shapes.find(s => s.id === 1).text, last + '\n\n');
    const revision = last === 'Hello' ? 0 : words.length;
    assert.equal(document.revision, revision);
    assert.equal(document.dirty, revision !== 0);
    if (revision === 0) assert.deepEqual(bytes, source);
    assert.throws(() => { shape.text = 'Too long\n\n'; });
    assert.deepEqual(Buffer.from(document.serialize()), bytes);
    assert.equal(document.revision, revision);
    const caseName = kind + '-' + (words.length > 1 ? 'repeat' : last);
    const outputPath = join(output, caseName + '.vsd'); writeFileSync(outputPath, bytes);
    cases.push({ caseName, sourcePath, sourceSha256: hash, outputPath, outputSha256: sha(bytes), revision,
      words, storedAfter: last + '\n\n', nativeExpectedAfter: last + '\n', unsupportedLengthRefusedAtomically: true });
  }
}
const modules = ['index.js','vsd-document.js','vsd-reader.js','vsd-writer.js','vsd-compression.js','vsd-compression-fit.js'].map(name => {
  const path = join(dirname(modulePath), name); return { path, sha256: sha(readFileSync(path)) };
});
const manifest = { modulePath, modules, cases, nativeValidatedByThisRun: false, fullFidelity: false };
writeFileSync(join(output, 'text-safety-generation.json'), JSON.stringify(manifest, null, 2));
console.log(JSON.stringify({ cases: cases.length, nativeValidatedByThisRun: false, manifest: join(output, 'text-safety-generation.json') }));
