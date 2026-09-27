import { cp, mkdir, rm, readFile, writeFile } from 'node:fs/promises';
import { createHash } from 'node:crypto';
import { fileURLToPath } from 'node:url';
import { generateSitePages } from './generate-site-pages.mjs';

const site = fileURLToPath(new URL('../.pages-site/', import.meta.url));
// `site` is the explicit repository-local generated Pages directory.
await rm(site, { recursive: true, force: true });
await mkdir(site, { recursive: true });
await cp(new URL('../docs/site.css', import.meta.url), `${site}site.css`);
await cp(new URL('../docs/theme.js', import.meta.url), `${site}theme.js`);
await cp(new URL('../demo/', import.meta.url), new URL('../.pages-site/demo/', import.meta.url), {
	recursive: true,
});
await mkdir(new URL('../.pages-site/demo/vendor/ole2/', import.meta.url), { recursive: true });
const samples = Bun.spawnSync(
	['bun', 'scripts/generate-demo-samples.mjs', '--output-dir', `${site}demo/samples`],
	{ cwd: fileURLToPath(new URL('../', import.meta.url)), stdout: 'inherit', stderr: 'inherit' },
);
if (samples.exitCode !== 0) throw new Error('Demo sample generation failed.');
const pages = await generateSitePages(site);
const result = await Bun.build({
	entrypoints: [fileURLToPath(new URL('../src/index.ts', import.meta.url))],
	target: 'browser',
	format: 'esm',
	outdir: `${site}demo/vendor/ole2`,
	naming: 'index.js',
});
if (!result.success) throw new AggregateError(result.logs, 'Browser demo bundle failed');
// Version the demo entry so returning visitors receive updated controls and samples.
const demoEntry = await readFile(`${site}demo/main.js`);
const demoVersion = createHash('sha256').update(demoEntry).digest('hex').slice(0, 12);
const demoHtml = await readFile(`${site}demo/index.html`, 'utf8');
await writeFile(`${site}demo/index.html`, demoHtml.replace('src="./main.js"', `src="./main.js?v=${demoVersion}"`));
console.log(`Built documentation and browser demo in ${site} (${pages.pageCount} pages generated)`);
