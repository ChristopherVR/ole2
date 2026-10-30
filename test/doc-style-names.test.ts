import { readFileSync, readdirSync } from 'node:fs';
import { describe, expect, it } from 'vitest';
import { readDocModel } from '../src/index.js';
import { builtInStyleName } from '../src/doc-styles.js';

describe('built-in style names', () => {
	it('agree with every built-in style Word 16 wrote in the fixtures', () => {
		const dir = new URL('./fixtures/doc/', import.meta.url);
		let compared = 0;
		for (const file of readdirSync(dir).filter((f) => f.endsWith('.doc'))) {
			const m = readDocModel(new Uint8Array(readFileSync(new URL(file, dir))));
			for (const style of m?.styles ?? []) {
				const known = style && style.builtIn ? builtInStyleName(style.sti) : undefined;
				if (!style || known === undefined) continue;
				// Word stores the name with spaces for some styles and without for others.
				expect(known.replace(/ /g, '').toLowerCase(), `${file} sti ${style.sti}`).toBe(
					style.name.replace(/ /g, '').toLowerCase(),
				);
				compared++;
			}
		}
		expect(compared).toBeGreaterThan(20);
	});

	it('names the identifiers that are known and no others', () => {
		expect(builtInStyleName(62)).toBe('Title');
		expect(builtInStyleName(85)).toBe('Hyperlink');
		expect(builtInStyleName(9999)).toBeUndefined();
	});
});
