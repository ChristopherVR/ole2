import { readFile } from 'node:fs/promises';
import { spawnSync } from 'node:child_process';
import assert from 'node:assert/strict';
const pkg = JSON.parse(await readFile(new URL('../package.json', import.meta.url), 'utf8'));
const tag = process.env.RELEASE_TAG;
assert(
	tag === `v${pkg.version}`,
	`RELEASE_TAG must exactly match v${pkg.version}; got ${tag || '(missing)'}`,
);
assert(/^\d+\.\d+\.\d+(?:-[0-9A-Za-z.-]+)?$/.test(pkg.version), 'Invalid package version');
assert.equal(pkg.name, '@christophervr/ole2');
assert(!pkg.private, 'Package must be public');
for (const spec of Object.values(pkg.dependencies || {}))
	assert(!/^(?:file|link|workspace):/.test(spec), 'Local dependency leaked into release');
const result = spawnSync('git', ['rev-parse', `${tag}^{commit}`], { encoding: 'utf8' });
assert.equal(result.status, 0, 'Release tag must exist');
const head = spawnSync('git', ['rev-parse', 'HEAD'], { encoding: 'utf8' });
assert.equal(result.stdout.trim(), head.stdout.trim(), 'Checkout must match the release tag');
if (process.argv.includes('--check')) {
	console.log(`Validated ${pkg.name}@${pkg.version} from ${tag}`);
	process.exit(0);
}
const response = await fetch(
	`https://registry.npmjs.org/${encodeURIComponent(pkg.name)}/${pkg.version}`,
);
if (response.ok) {
	console.log(`${pkg.name}@${pkg.version} is already published; nothing to do.`);
	process.exit(0);
}
assert.equal(
	response.status,
	404,
	`Registry lookup failed (${response.status}); refusing to guess publication state`,
);
const args = ['publish', '--access', 'public', '--provenance'];
if (pkg.version.includes('-')) args.push('--tag', 'next');
// The CI runner is Linux; token auth can bootstrap, then npm uses the configured OIDC publisher.
const publish = spawnSync(process.platform === 'win32' ? 'npm.cmd' : 'npm', args, {
	stdio: 'inherit',
	shell: process.platform === 'win32',
});
process.exit(publish.status ?? 1);
