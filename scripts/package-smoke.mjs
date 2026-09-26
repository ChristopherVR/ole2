import { mkdtemp, readFile, writeFile, readdir } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join, resolve } from 'node:path';
import { spawnSync } from 'node:child_process';
import assert from 'node:assert/strict';
const directory = await mkdtemp(join(tmpdir(), 'ole2-package-'));
const npm = process.platform === 'win32' ? 'npm.cmd' : 'npm';
function run(command, args, cwd = process.cwd(), capture = false) {
 const result = spawnSync(command, args, { cwd, encoding: 'utf8', stdio: capture ? 'pipe' : 'inherit', shell: process.platform === 'win32' && command.endsWith('.cmd') });
 if (result.status !== 0) throw new Error(`${command} failed (${result.status}): ${result.stderr || ''}`);
 return result.stdout;
}
const packed = JSON.parse(run(npm, ['pack', '--ignore-scripts', '--json', '--pack-destination', directory], undefined, true))[0];
assert(packed.files.some(file => file.path === 'dist/index.js'));
assert(!packed.files.some(file => file.path.includes('docx')), 'Modern DOCX code must not ship in the legacy package');
assert(!packed.files.some(file => /^(?:src|test|node_modules)\//.test(file.path)));
const tarball = join(directory, packed.filename);
await writeFile(join(directory, 'package.json'), JSON.stringify({ name: 'ole2-consumer-smoke', private: true, type: 'module' }));
run(npm, ['install', '--ignore-scripts', '--no-audit', '--no-fund', tarball], directory);
const installed = JSON.parse(await readFile(join(directory, 'node_modules/@christophervr/ole2/package.json'), 'utf8'));
assert.equal(installed.name, '@christophervr/ole2');
await writeFile(join(directory, 'verify.mjs'), `
import assert from 'node:assert/strict';
import { buildOle2, parseOle2 } from '@christophervr/ole2';

const bytes = buildOle2(new Map([['Sample', new Uint8Array([1,2,3])]]));
assert.deepEqual([...parseOle2(bytes).getStream('Sample')], [1,2,3]);



console.log('Packed legacy OLE2 package works in an independent npm consumer.');
`);
run(process.execPath, [join(directory, 'verify.mjs')], directory);
console.log(`Verified ${packed.filename}; isolated consumer: ${directory}`);

