import { cp, mkdir, rm, readFile, writeFile } from 'node:fs/promises';
import { createHash } from 'node:crypto';
import { join, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import { generateSitePages } from './generate-site-pages.mjs';

const repo = fileURLToPath(new URL('../', import.meta.url));
const generatedPages = resolve(repo, '.pages-site');
const checkPages = resolve(repo, '.pages-check');
const site = process.env.OLE2_PAGES_OUTPUT_DIR
	? resolve(process.env.OLE2_PAGES_OUTPUT_DIR)
	: generatedPages;
if (site !== generatedPages && site !== checkPages) {
	throw new Error('OLE2_PAGES_OUTPUT_DIR must be .pages-site or .pages-check within this repository.');
}
const includeDemo = process.env.OLE2_BUILD_PAGES_DEMO === '1';
const siteFile = (...parts) => join(site, ...parts);

// Remove stale demo routes when the default, demo-free Pages build runs.
await rm(site, { recursive: true, force: true });
await mkdir(site, { recursive: true });
await cp(new URL('../docs/site.css', import.meta.url), siteFile('site.css'));
await cp(new URL('../docs/theme.js', import.meta.url), siteFile('theme.js'));
if (includeDemo) {
	await cp(new URL('../demo/', import.meta.url), siteFile('demo'), { recursive: true });
}
const pages = await generateSitePages(site);

if (includeDemo) {
	await mkdir(siteFile('demo', 'vendor', 'ole2'), { recursive: true });
	const samples = Bun.spawnSync(
		['bun', 'scripts/generate-demo-samples.mjs', '--output-dir', siteFile('demo', 'samples')],
		{ cwd: fileURLToPath(new URL('../', import.meta.url)), stdout: 'inherit', stderr: 'inherit' },
	);
	if (samples.exitCode !== 0) throw new Error('Demo sample generation failed.');
	const result = await Bun.build({
		entrypoints: [fileURLToPath(new URL('../src/index.ts', import.meta.url))],
		target: 'browser',
		format: 'esm',
		outdir: siteFile('demo', 'vendor', 'ole2'),
		naming: 'index.js',
	});
	if (!result.success) throw new AggregateError(result.logs, 'Browser demo bundle failed');
	// Version the demo entry so returning visitors receive updated controls and samples.
	const demoEntry = await readFile(siteFile('demo', 'main.js'));
	const demoVersion = createHash('sha256').update(demoEntry).digest('hex').slice(0, 12);
	const demoHtml = await readFile(siteFile('demo', 'index.html'), 'utf8');
	await writeFile(siteFile('demo', 'index.html'), demoHtml.replace('src="./main.js"', `src="./main.js?v=${demoVersion}"`));
}

console.log(
	`Built documentation${includeDemo ? ' and browser demo' : ''} in ${site} (${pages.pageCount} pages generated)`,
);
