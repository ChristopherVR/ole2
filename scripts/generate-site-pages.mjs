import { access, mkdir, readFile, writeFile } from 'node:fs/promises';
import { resolve } from 'node:path';
import { pathToFileURL, fileURLToPath } from 'node:url';
import { apiPages } from '../docs/api-pages.mjs';
import { highlightCodeBlocks } from './highlight-code.mjs';

const templatePath = new URL('../docs/site-template.html', import.meta.url);
const homeContentPath = new URL('../docs/home-content.html', import.meta.url);
const escapeHtml = (value) => String(value).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;').replace(/'/g, '&#39;');
const inline = (value) => escapeHtml(value).replace(/`([^`]+)`/g, '<code>$1</code>');

function pageContent(page) {
	const sections = page.sections.map((section) => `
		<section class="guide-section">
			<div class="guide-section__copy"><p class="eyebrow">IMPLEMENTATION NOTES</p><h2>${escapeHtml(section.title)}</h2><p>${inline(section.body)}</p></div>
			<ul class="guide-points">${section.bullets.map((item) => `<li>${inline(item)}</li>`).join('')}</ul>
		</section>`).join('');
	return `
		<nav class="breadcrumbs" aria-label="Breadcrumb"><a href="../../">Home</a><span aria-hidden="true">/</span><span>API guide</span></nav>
		<section class="guide-hero" aria-labelledby="page-title"><p class="eyebrow">${escapeHtml(page.kicker)}</p><h1 id="page-title">${escapeHtml(page.title)}<br /><em>API guide</em></h1><p class="lede">${escapeHtml(page.intro)}</p></section>
		${sections}
		<section class="example-section"><div class="section-heading"><div><p class="eyebrow">QUICK EXAMPLE</p><h2>Use the supported API</h2></div><p>Import the package root unless a subpath is shown in the sample.</p></div><div class="highlighted-code"><div class="code-label">${escapeHtml(page.title)} · TypeScript</div><pre><code>${escapeHtml(page.example)}</code></pre></div></section>
		<section class="guide-foot"><p>This guide documents the current package API. Unsupported or unsafe operations are described explicitly; format detection alone does not imply content support.</p><a class="button button--ghost" href="../../">Back to format overview</a></section>`;
}

export async function generateSitePages(outputDirectory) {
	const output = outputDirectory instanceof URL ? fileURLToPath(outputDirectory) : outputDirectory;
	const template = await readFile(templatePath, 'utf8');
	const homeHtml = await readFile(homeContentPath, 'utf8');
	const paths = Object.fromEntries(apiPages.map((page) => [page.slug, `{{base}}api/${page.slug}/`]));
	const links = apiPages.map((page) => `<a href="${paths[page.slug]}">${escapeHtml(page.title.replace('Legacy ', ''))}</a>`).join('');
	const demoEnabled = process.env.OLE2_BUILD_PAGES_DEMO === '1';
	const render = async (content, title, description, base) => {
		const demoUrl = `${base}demo/`;
		const demoAction = demoEnabled ? `<a class="button button--ghost" href="${demoUrl}">Try the browser demo</a>` : '';
		const demoNav = demoEnabled ? `<a href="${demoUrl}">Live demo</a>` : '';
		let html = template.replace('{{content}}', content).replaceAll('{{api_links}}', links).replaceAll('{{demo_action}}', demoAction).replaceAll('{{demo_nav}}', demoNav).replaceAll('{{title}}', escapeHtml(title)).replaceAll('{{description}}', escapeHtml(description)).replaceAll('{{base}}', base).replaceAll('{{home}}', `${base}`);
		for (const page of apiPages) html = html.replaceAll(`{{${page.slug}}}`, `${base}api/${page.slug}/`);
		html = html.replaceAll('{{shared}}', `${base}api/containers/`);
		const unresolved = html.match(/\{\{[^}]+\}\}/g);
		if (unresolved) throw new Error(`Unresolved template values in ${title}: ${[...new Set(unresolved)].join(', ')}`);
		return (await highlightCodeBlocks(html)).html;
	};
	await mkdir(output, { recursive: true });
	const home = await render(homeHtml, 'Legacy Office API guide', 'Bounded OLE2 and legacy Office codecs, with documented format support and limitations.', './');
	await writeFile(`${output}/index.html`, home);
	for (const page of apiPages) {
		const directory = `${output}/api/${page.slug}`;
		await mkdir(directory, { recursive: true });
		const html = await render(pageContent(page), page.title, page.description, '../../');
		await writeFile(`${directory}/index.html`, html);
	}
	const htmlFiles = ['index.html', ...apiPages.map((page) => `api/${page.slug}/index.html`)];
	for (const file of htmlFiles) {
		const source = await readFile(`${output}/${file}`, 'utf8');
		const sourceUrl = pathToFileURL(resolve(output, file));
		for (const [, link] of source.matchAll(/(?:href|src)="([^"]+)"/g)) {
			if (/^(?:https?:|mailto:|#|data:)/i.test(link)) continue;
			const target = new URL(link, sourceUrl);
			const targetPath = fileURLToPath(target);
			const candidate = target.pathname.endsWith('/') ? resolve(targetPath, 'index.html') : targetPath;
			try { await access(candidate); } catch { throw new Error(`Broken local link in ${file}: ${link}`); }
		}
	}
	return { pageCount: apiPages.length + 1 };
}

if (process.argv[1] && fileURLToPath(import.meta.url) === process.argv[1]) {
	const output = process.argv[2] ?? fileURLToPath(new URL('../.pages-site', import.meta.url));
	const result = await generateSitePages(output);
	console.log(`Generated ${result.pageCount} Pages documents in ${output}`);
}
