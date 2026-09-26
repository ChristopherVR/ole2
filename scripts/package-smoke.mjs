import { mkdtemp, readFile, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join, dirname } from 'node:path';
import { spawnSync } from 'node:child_process';
import assert from 'node:assert/strict';
const directory = await mkdtemp(join(tmpdir(), 'ole2-package-'));
const npm = process.platform === 'win32' ? process.execPath : 'npm';
const npmArgs =
	process.platform === 'win32'
		? [join(dirname(process.execPath), 'node_modules/npm/bin/npm-cli.js')]
		: [];
function run(command, args, cwd = process.cwd(), capture = false) {
	const result = spawnSync(command, args, {
		cwd,
		encoding: 'utf8',
		stdio: capture ? 'pipe' : 'inherit',
	});
	if (result.error) throw result.error;
	if (result.status !== 0)
		throw new Error(`${command} failed (${result.status}): ${result.stderr || ''}`);
	return result.stdout;
}
const packed = JSON.parse(
	run(
		npm,
		[...npmArgs, 'pack', '--ignore-scripts', '--json', '--pack-destination', directory],
		undefined,
		true,
	),
)[0];
assert(packed.files.some((file) => file.path === 'dist/index.js'));
assert(
	!packed.files.some((file) => file.path.includes('docx')),
	'Modern DOCX code must not ship in the legacy package',
);
assert(!packed.files.some((file) => /^(?:src|test|node_modules)\//.test(file.path)));
const tarball = join(directory, packed.filename);
await writeFile(
	join(directory, 'package.json'),
	JSON.stringify({ name: 'ole2-consumer-smoke', private: true, type: 'module' }),
);
run(
	npm,
	[...npmArgs, 'install', '--ignore-scripts', '--no-audit', '--no-fund', tarball],
	directory,
);
const installed = JSON.parse(
	await readFile(join(directory, 'node_modules/@christophervr/ole2/package.json'), 'utf8'),
);
assert.equal(installed.name, '@christophervr/ole2');
await writeFile(
	join(directory, 'verify.mjs'),
	`
import assert from 'node:assert/strict';
import { buildOle2, parseOle2, readCompoundFileStream, replaceCompoundFileStream, readOleXlsGrid, inspectLegacyVisio, inspectLegacyPublisher, writeLegacyOfficeMetadata } from '@christophervr/ole2';
import { buildPptFile } from '@christophervr/ole2/legacy-ppt-writer';
import { readRecord } from '@christophervr/ole2/legacy-ppt-record-stream';

const bytes = buildOle2(new Map([['Sample', new Uint8Array([1,2,3])]]));
assert.deepEqual([...parseOle2(bytes).getStream('Sample')], [1,2,3]);
const changed = replaceCompoundFileStream(new Uint8Array(bytes), ['Sample'], new Uint8Array([3,2,1]));
assert.deepEqual([...readCompoundFileStream(changed, ['Sample'])], [3,2,1]);
for (const api of [readOleXlsGrid, inspectLegacyVisio, inspectLegacyPublisher, writeLegacyOfficeMetadata, readRecord]) assert.equal(typeof api, 'function');
const ppt = await buildPptFile({ widthEmu: 9144000, heightEmu: 5143500, slides: [{ shapes: [] }], pictures: [] });
assert.ok(parseOle2(ppt.buffer).getStream('PowerPoint Document')?.length > 0);
console.log('Packed legacy OLE2 package works in an independent npm consumer.');
`,
);
run(process.execPath, [join(directory, 'verify.mjs')], directory);
console.log(`Verified ${packed.filename}; isolated consumer: ${directory}`);
