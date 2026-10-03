import assert from 'node:assert/strict';
import { mkdtemp, mkdir, readFile, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { highlightCodeBlocks } from './highlight-code.mjs';
import { generateSitePages } from './generate-site-pages.mjs';

const typescript = 'const answer: number = 42;\nif (answer < 50) console.log("<ready>");';
const shell = 'bun run build';
const input = `<main><pre><code>${typescript.replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;')}</code></pre><pre><code>${shell}</code></pre></main>`;
const result = await highlightCodeBlocks(input);
assert.equal(result.count, 2);
assert.match(result.html, /class="shiki github-dark"/);
assert.match(result.html, /data-language="typescript"/);
assert.match(result.html, /data-language="bash"/);
assert.match(result.html, /data-copy-source="const answer: number = 42;\nif \(answer &lt; 50\) console\.log\(&quot;&lt;ready&gt;&quot;\);"/);
assert.match(result.html, /class="copy-code"/);

const site = await mkdtemp(join(tmpdir(), 'ole2-pages-'));
try {
	await mkdir(join(site, 'demo'), { recursive: true });
	await writeFile(join(site, 'demo/index.html'), '<main>demo</main>');
	await writeFile(join(site, 'site.css'), '');
	await writeFile(join(site, 'theme.js'), '');
	const generated = await generateSitePages(site);
	assert.equal(generated.pageCount, 8);
	const home = await readFile(join(site, 'index.html'), 'utf8');
	assert.match(home, /href="\.\/api\/word\/"/);
	assert.match(home, /href="\.\/api\/models\/"/);
	assert.match(home, /href="\.\/api\/publisher\/"/);
	assert.doesNotMatch(home, /\{\{[^}]+\}\}/);
	const word = await readFile(join(site, 'api/word/index.html'), 'utf8');
	assert.match(word, /class="shiki github-dark"/);
	assert.match(word, /href="\.\.\/\.\.\/api\/excel\/"/);
	const visio = await readFile(join(site, 'api/visio/index.html'), 'utf8');
	assert.match(visio, /TrailerStream pointer/);
	assert.match(visio, /not decoded or editable/i);
	assert.match(visio, /data-language="typescript"/);
	assert.doesNotMatch(visio, /\{\{[^}]+\}\}/);
} finally {
	await rm(site, { recursive: true, force: true });
}
console.log('Page highlighting, guide generation, relative links, and copy-source checks passed.');
