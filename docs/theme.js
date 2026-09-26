const root = document.documentElement;
const preferenceKey = 'ole2-site-theme';

let selected;
try {
	selected = localStorage.getItem(preferenceKey);
} catch {
	selected = null;
}

const initial = selected === 'light' || selected === 'dark'
	? selected
	: matchMedia('(prefers-color-scheme: dark)').matches ? 'dark' : 'light';
root.dataset.theme = initial;

for (const toggle of document.querySelectorAll('[data-theme-toggle]')) {
	const update = (theme) => {
		const dark = theme === 'dark';
		toggle.textContent = dark ? '☼' : '◐';
		toggle.setAttribute('aria-label', dark ? 'Switch to light theme' : 'Switch to dark theme');
		toggle.setAttribute('aria-pressed', String(dark));
		toggle.title = dark ? 'Switch to light theme' : 'Switch to dark theme';
	};
	update(initial);
	toggle.addEventListener('click', () => {
		const theme = root.dataset.theme === 'dark' ? 'light' : 'dark';
		root.dataset.theme = theme;
		try {
			localStorage.setItem(preferenceKey, theme);
		} catch {
			// Theme switching still works for this page when storage is unavailable.
		}
		update(theme);
	});
}
