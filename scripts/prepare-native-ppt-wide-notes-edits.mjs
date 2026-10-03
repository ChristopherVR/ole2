import { readFileSync, writeFileSync, mkdirSync } from 'node:fs';
import { createHash } from 'node:crypto';
import { join, resolve } from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';
// Requires the committed repository-owned wide-notes fixture and a package build.
// Generation is separate from native read-only/saved-control snapshot evidence.
const output = resolve(process.argv[2] ?? '.native-validation/ppt-notes-wide');
const dist = resolve(process.argv[3] ?? 'dist');
const { PptDocument } = await import(pathToFileURL(join(dist, 'ppt-document.js')).href);
const source = fileURLToPath(new URL('../test/fixtures/ppt/wide-notes.ppt', import.meta.url));
const input = new Uint8Array(readFileSync(source)), hash = bytes => createHash('sha256').update(bytes).digest('hex');
const sourceSha256 = '86905ce486c3fae0d4d63bc3f8187ca3d2ac87b46a599a7c6963d337b189ad42';
if (hash(input) !== sourceSha256) throw new Error('Owned wide notes fixture changed; new native evidence required');
const document = new PptDocument(input), notes = document.slides[1].notes;
const bodies = notes?.texts.filter(text => text.role === 'body');
if (bodies?.length !== 1 || bodies[0].encoding !== 'utf16') throw new Error('Expected one existing UTF16 notes body slot');
const before = '\u6771\u4eac \u03a9 \ud83d\ude00\rWide notes.', after = '\u5927\u962a \u03b2 \ud83d\ude00\rWide model.';
if (bodies[0].text !== before) throw new Error('Wide notes source text differs');
bodies[0].text = after;
const bytes = document.serialize(), outputSha256 = hash(bytes);
if (outputSha256 !== '01abcbfe1bdf12afa02cbf9cd1e61a661d8d489b56422b2689d7f25e078c8b20') throw new Error('Wide notes edit bytes changed; new native evidence required');
mkdirSync(output, { recursive: true });
const path = join(output, 'wide-notes-edited.ppt'), changes = [{ path: '/slides/1/notes/0', value: after }];
writeFileSync(path, bytes); writeFileSync(join(output, 'wide-notes-changes.json'), JSON.stringify(changes, null, 2));
writeFileSync(join(output, 'wide-notes-generation.json'), JSON.stringify({ schemaVersion: 1, source, sourceSha256, output: path, outputSha256,
  provenance: 'Repository-owned native derivative; generator and source hash pinned in fixture manifest',
  notesId: notes.notesId, notesPersistId: notes.persistId, shapeId: bodies[0].shapeId, before, after,
  serialization: 'PptDocument.serialize()', changes, nativeValidatedByThisRun: false, fullFidelity: false }, null, 2));
console.log(JSON.stringify({ output: path, outputSha256, nativeValidatedByThisRun: false }));
