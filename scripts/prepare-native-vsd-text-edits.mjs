import { readFileSync, writeFileSync, mkdirSync } from 'node:fs';
import { createHash } from 'node:crypto';
import { resolve, dirname, join } from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';
import assert from 'node:assert/strict';
// Generate only an owned neutral fixture edit. This does not launch Office;
// native snapshots and save/reopen comparisons are a separate evidence gate.
const root = resolve(dirname(fileURLToPath(import.meta.url)), '..');
const modulePath = resolve(process.argv[2] ?? join(root, 'dist/index.js'));
const output = resolve(process.argv[3] ?? join(root, '.native-validation/vsd-text'));
const { parseVsd } = await import(pathToFileURL(modulePath).href);
const sourcePath = join(root, 'test/fixtures/vsd/native-visio16-v11.vsd');
const source = readFileSync(sourcePath);
const sha = bytes => createHash('sha256').update(bytes).digest('hex');
assert.equal(sha(source), 'c6c97822e7bb2cc3e96da9d7d35fe74ac16a7f2c90d19ad4c80967d2896e0974');
const document = parseVsd(source), shape = document.pages[0].shapes.find(s => s.id === 1);
assert.equal(shape.text, 'Hello\n\n');
shape.text = shape.text;
assert.equal(document.revision, 0);
assert.equal(document.dirty, false);
assert.deepEqual(Buffer.from(document.serialize()), source);
shape.text = 'World\n\n';
const bytes = Buffer.from(document.serialize());
assert.equal(sha(bytes), 'cc257f441f2d30fb30ea9b0281fcf4570c092532518a4894a28fee4ae9bf3bdb');
assert.equal(parseVsd(bytes).pages[0].shapes.find(s => s.id === 1).text, 'World\n\n');
assert.equal(document.revision, 1);
assert.equal(document.dirty, true);
assert.throws(() => { shape.text = 'Too long\n\n'; });
assert.deepEqual(Buffer.from(document.serialize()), bytes);
assert.equal(document.revision, 1);
mkdirSync(output, { recursive: true });
const path = join(output, 'native-model-world.vsd');
writeFileSync(path, bytes);
const modules = ['index.js', 'vsd-document.js', 'vsd-reader.js', 'vsd-writer.js', 'vsd-compression.js', 'vsd-compression-fit.js'].map(name => {
  const path = join(dirname(modulePath), name); return { path, sha256: sha(readFileSync(path)) };
});
const manifest = { sourcePath, sourceSha256: sha(source), modulePath, modules,
  output: path, outputSha256: sha(bytes), revision: document.revision, dirty: document.dirty,
  storedBefore: 'Hello\n\n', storedAfter: 'World\n\n', nativeConsumerBefore: 'Hello\n', nativeConsumerAfter: 'World\n',
  nativeValidatedByThisRun: false, fullFidelity: false };
writeFileSync(join(output, 'text-generation.json'), JSON.stringify(manifest, null, 2));
console.log(JSON.stringify({ output: path, outputSha256: sha(bytes), nativeValidatedByThisRun: false }));
