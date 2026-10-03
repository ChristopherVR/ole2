import { readFileSync, writeFileSync, mkdirSync } from 'node:fs';
import { createHash } from 'node:crypto';
import { join, resolve } from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';

// Build first. Explicit repository-owned inputs only; never launch Office here.
const output = resolve(process.argv[2] ?? '.native-validation/ppt-notes');
const dist = resolve(process.argv[3] ?? 'dist');
const { PptDocument } = await import(pathToFileURL(join(dist, 'ppt-document.js')).href);
const source = fileURLToPath(new URL('../test/fixtures/ppt/native-text.ppt', import.meta.url));
const input = new Uint8Array(readFileSync(source));
const hash = bytes => createHash('sha256').update(bytes).digest('hex');
const sourceSha256 = '2bae7c1501262b79d2d9840ba968a5ed7b4ee12facde52082d3b114772f9b4e0';
if (hash(input) !== sourceSha256) throw new Error('Owned fixture hash changed; new native evidence required');
const document = new PptDocument(input);
const replacements = ['Synthetic notes modified.', 'Unicode notes modified.'];
const identities = [];
for (let index = 0; index < replacements.length; index++) {
  const notes = document.slides[index].notes, bodies = notes?.texts.filter(text => text.role === 'body');
  if (bodies?.length !== 1) throw new Error('Expected one validated body slot per owned notes page');
  const text = bodies[0];
  identities.push({ slideId: notes.slideId, slidePersistId: notes.slidePersistId, notesId: notes.notesId,
    notesPersistId: notes.persistId, shapeId: text.shapeId, encoding: text.encoding, before: text.text, after: replacements[index] });
  text.text = replacements[index];
}
const bytes = document.serialize(), outputSha256 = hash(bytes);
if (outputSha256 !== 'ea4de78e3e5671f38a8cbcd6c8845d76d2ec8d9cdf6f1d6f65b00b29f4983643') throw new Error('Edited bytes changed; obtain new native evidence');
if (document.revision !== 2) throw new Error('Expected two typed-model transactions');
mkdirSync(output, { recursive: true });
const path = join(output, 'notes-edited.ppt');
writeFileSync(path, bytes);
const changes = replacements.map((value, index) => ({ path: `/slides/${index}/notes/0`, value }));
writeFileSync(join(output, 'notes-changes.json'), JSON.stringify(changes, null, 2));
writeFileSync(join(output, 'notes-generation.json'), JSON.stringify({ schemaVersion: 1, source, sourceSha256, output: path, outputSha256,
  provenance: 'Repository-owned PowerPoint-created synthetic fixture; see test/fixtures/ppt/README.md and fixture manifest',
  serialization: 'PptDocument.serialize()', identities, changes,
  changedBytes: input.reduce((count, byte, index) => count + Number(byte !== bytes[index]), 0),
  evidence: 'generation-only; independent Office snapshots and save/reopen comparison required',
  nativeValidatedByThisRun: false, fullFidelity: false }, null, 2));
console.log(JSON.stringify({ generated: path, outputSha256, nativeValidatedByThisRun: false }));
