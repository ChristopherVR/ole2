import { readFileSync, readdirSync, statSync, writeFileSync } from 'node:fs';
import { createHash } from 'node:crypto';
import { spawnSync } from 'node:child_process';
import { resolve, join } from 'node:path';

// Independent parser oracle, not a Visio native or render-fidelity assertion.
// Explicit fixture input only; invokes a trusted, separately supplied vsd2raw.
const [mode, tool, input, output, distro = 'Ubuntu'] = process.argv.slice(2);
if (!['--native', '--wsl'].includes(mode) || !tool || !input || !output) {
  throw new Error('Usage: node scripts/libvisio-snapshot.mjs --native VSD2RAW.exe INPUT.vsd OUTPUT.json\nOr: --wsl EXTRACTED_PACKAGE_ROOT INPUT.vsd OUTPUT.json [DISTRO]');
}
const inputPath = resolve(input);
if (statSync(inputPath).size > 16 * 1024 * 1024) throw new Error('Fixture exceeds 16 MiB input limit');
const toWsl = p => {
  const match = /^([A-Za-z]):[\\/](.*)$/.exec(resolve(p));
  if (!match) throw new Error('WSL mode requires explicit Windows drive paths');
  return `/mnt/${match[1].toLowerCase()}/${match[2].replaceAll('\\', '/')}`;
};
let command, prefix, executable, libraryDirectory;
if (mode === '--native') {
  executable = resolve(tool);
  command = executable;
  prefix = [];
} else {
  const root = resolve(tool);
  libraryDirectory = join(root, 'usr', 'lib', 'x86_64-linux-gnu');
  executable = join(root, 'usr', 'bin', 'vsd2raw');
  command = 'wsl.exe';
  prefix = ['-d', distro, '--', 'env', `LD_LIBRARY_PATH=${toWsl(join(root, 'usr', 'lib', 'x86_64-linux-gnu'))}`, toWsl(executable)];
}
const run = args => {
  const result = spawnSync(command, [...prefix, ...args], { encoding: 'utf8', timeout: 30000, maxBuffer: 8 * 1024 * 1024, windowsHide: true });
  if (result.error) throw result.error;
  if (result.status !== 0) throw new Error(`libvisio refused input (${result.status}): ${result.stderr.slice(0, 2048)}`);
  return result;
};
const version = run(['--version']).stdout.trim();
const parsed = run([mode === '--wsl' ? toWsl(inputPath) : inputPath]);
const callbacks = parsed.stdout.replaceAll('\r\n', '\n');
if (!callbacks.includes('startDocument()') || !/^endDocument(?:\(\))?\s*$/m.test(callbacks) || !callbacks.includes('startPage(')) {
  throw new Error('Reader did not produce a complete nonempty drawing callback document');
}
if (/:\s*(?:nan|[+-]?inf)(?:in|pt|cm|%)(?:\W|$)/i.test(callbacks)) {
  throw new Error('Reader emitted non-finite drawing coordinates; semantic acceptance not established');
}
const sha256 = p => createHash('sha256').update(readFileSync(p)).digest('hex');
const parserLibraries = mode === '--wsl' ? readdirSync(libraryDirectory)
  .filter(name => /^lib(?:visio|revenge).*\.so\.\d+\.\d+\.\d+$/.test(name))
  .map(name => { const p = join(libraryDirectory, name); return { path: p, sha256: sha256(p) }; }) : [];
writeFileSync(output, JSON.stringify({
  consumer: 'libvisio vsd2raw', version, execution: mode === '--wsl' ? `WSL ${distro}` : 'native executable',
  evidence: 'independent-library-drawing-callbacks', nativeVisio: false, fullFidelity: false,
  input: { path: inputPath, sha256: sha256(inputPath) }, executable: { path: executable, sha256: sha256(executable) },
  parserLibraries, libraryProvenanceScope: 'WSL extracted parser/render-callback libraries only; complete dynamic dependency inventory not asserted',
  limits: { inputBytes: 16 * 1024 * 1024, outputBytes: 8 * 1024 * 1024, timeoutMilliseconds: 30000 },
  pages: (callbacks.match(/startPage\(/g) ?? []).length,
  paths: (callbacks.match(/drawPath\s*\(/g) ?? []).length,
  textObjects: (callbacks.match(/startTextObject\s*\(/g) ?? []).length,
  diagnostics: parsed.stderr, callbacks,
}, null, 2) + '\n');
console.log(JSON.stringify({ version, output: resolve(output), nativeVisio: false }));
