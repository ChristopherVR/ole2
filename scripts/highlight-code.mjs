import { createHighlighter } from 'shiki';

const highlighter = await createHighlighter({
	themes: ['github-dark'],
	langs: ['typescript', 'javascript', 'bash'],
});

function decodeHtml(value) {
	return value
		.replace(/&lt;/g, '<')
		.replace(/&gt;/g, '>')
		.replace(/&quot;/g, '"')
		.replace(/&#39;/g, "'")
		.replace(/&amp;/g, '&');
}

function escapeAttribute(value) {
	return value.replace(/&/g, '&amp;').replace(/"/g, '&quot;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}

function languageFor(source, explicit = '') {
	if (/^(?:npm|bun|node|npx|bunx)\s/m.test(source.trim())) return 'bash';
	return explicit === 'javascript' ? 'javascript' : 'typescript';
}

export async function highlightCodeBlocks(html) {
	const pattern = /<pre([^>]*)>\s*<code([^>]*)>([\s\S]*?)<\/code>\s*<\/pre>/gi;
	let index = 0;
	let transformed = '';
	let previous = 0;
	for (const match of html.matchAll(pattern)) {
		const full = match[0];
		const at = match.index;
		transformed += html.slice(previous, at);
		const rawSource = decodeHtml(match[3]);
		const source = rawSource.replace(/^\n/, '').replace(/\n$/, '');
		const langMatch = /language-([\w-]+)/i.exec(match[2]);
		const lang = languageFor(source, langMatch?.[1]?.toLowerCase());
		const shikiLanguage = lang === 'bash' ? 'bash' : lang;
		const highlighted = highlighter.codeToHtml(source, { lang: shikiLanguage, theme: 'github-dark' });
		const copy = escapeAttribute(source);
		transformed += `<div class="highlighted-code" data-language="${lang}"><div class="code-toolbar"><span>${lang === 'bash' ? 'Shell' : lang === 'javascript' ? 'JavaScript' : 'TypeScript'}</span><button class="copy-code" type="button" data-copy-source="${copy}" aria-label="Copy code">Copy</button></div>${highlighted}</div>`;
		previous = at + full.length;
		index++;
	}
	transformed += html.slice(previous);
	return { html: transformed, count: index };
}
