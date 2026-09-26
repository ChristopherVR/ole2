import { cp, mkdir } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';

const site = fileURLToPath(new URL('../.pages-site/', import.meta.url));
await mkdir(site, { recursive: true });
await cp(new URL('../docs/', import.meta.url), site, { recursive: true });
await cp(new URL('../demo/', import.meta.url), new URL('../.pages-site/demo/', import.meta.url), {
	recursive: true,
});
await mkdir(new URL('../.pages-site/demo/vendor/ole2/', import.meta.url), { recursive: true });
const result = await Bun.build({
	entrypoints: [fileURLToPath(new URL('../src/index.ts', import.meta.url))],
	target: 'browser',
	format: 'esm',
	outdir: `${site}demo/vendor/ole2`,
	naming: 'index.js',
});
if (!result.success) throw new AggregateError(result.logs, 'Browser demo bundle failed');
console.log(`Built documentation and browser demo in ${site}`);
