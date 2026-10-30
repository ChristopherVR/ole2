/** Field instruction tokenizing and typed decoding for the common field kinds
 * ([MS-DOC] 2.9.88 Fld, ECMA-376 17.16.5 field codes). */
import type { DocField, DocFieldBase } from './doc-objects-types.js';

export interface ParsedInstruction {
	readonly name: string;
	readonly args: readonly string[];
	readonly switches: Record<string, string | true>;
}

const ARG_SWITCHES: Readonly<Record<string, ReadonlySet<string>>> = {
	HYPERLINK: new Set(['o', 'l', 't']),
	TOC: new Set(['o', 't', 'b', 'c', 'f', 'l', 'p', 's', 'd', 'a']),
	REF: new Set(['d']), PAGEREF: new Set(['d']),
	SEQ: new Set(['r', 's']),
	DATE: new Set(), TIME: new Set(),
};
const FORMAT_SWITCHES = new Set(['*', '@', '#']);

/** Split an instruction into name, positional arguments and backslash switches. */
export function parseFieldInstruction(text: string): ParsedInstruction {
	const tokens: { text: string; quoted: boolean }[] = [];
	for (let i = 0; i < text.length;) {
		const c = text[i]!;
		if (c === ' ' || c === '\t' || c === '\r' || c === '\n' || c === ' ') { i++; continue; }
		if (c === '"' || c === '“') {
			let s = '';
			for (i++; i < text.length && text[i] !== '"' && text[i] !== '”'; i++) {
				if (text[i] === '\\' && (text[i + 1] === '\\' || text[i + 1] === '"')) i++;
				s += text[i];
			}
			i++;
			tokens.push({ text: s, quoted: true });
			continue;
		}
		let s = '';
		while (i < text.length && !/\s/u.test(text[i]!)) s += text[i++];
		tokens.push({ text: s, quoted: false });
	}
	const name = (tokens.shift()?.text ?? '').toUpperCase();
	const takes = ARG_SWITCHES[name];
	const args: string[] = [], switches: Record<string, string | true> = {};
	for (let i = 0; i < tokens.length; i++) {
		const t = tokens[i]!;
		if (!t.quoted && t.text.startsWith('\\') && t.text.length >= 2) {
			const key = t.text.slice(1, 2);
			const inline = t.text.slice(2);
			const next = tokens[i + 1];
			const wantsArg = FORMAT_SWITCHES.has(key) || takes?.has(key) === true;
			let value: string | true = true;
			if (inline) value = inline;
			else if (wantsArg && next && (next.quoted || !next.text.startsWith('\\'))) { value = next.text; i++; }
			const prior = switches[key];
			switches[key] = typeof prior === 'string' && typeof value === 'string' ? `${prior},${value}` : value;
		} else args.push(t.text);
	}
	return { name, args, switches };
}

type Base = Omit<DocFieldBase, 'kind'>;

function str(v: string | true | undefined): string | undefined {
	return typeof v === 'string' ? v : undefined;
}

/** Add the typed members for the instruction's field name. */
export function typeField(base: Base, parsed: ParsedInstruction): DocField {
	const sw = parsed.switches, arg0 = parsed.args[0];
	const fields = { ...base, switches: sw };
	switch (parsed.name) {
		case 'HYPERLINK':
			return { ...fields, kind: 'hyperlink', url: arg0, anchor: str(sw.l), tooltip: str(sw.o), target: str(sw.t) };
		case 'REF': case 'PAGEREF':
			return { ...fields, kind: parsed.name === 'REF' ? 'ref' : 'pageref', bookmark: arg0 ?? '' };
		case 'TOC':
			return { ...fields, kind: 'toc', outlineLevels: str(sw.o), hyperlinks: sw.h === true, tcEntries: sw.f !== undefined };
		case 'DATE': case 'TIME':
			return { ...fields, kind: parsed.name === 'DATE' ? 'date' : 'time', format: str(sw['@']) };
		case 'SEQ':
			return { ...fields, kind: 'seq', identifier: arg0 ?? '', reset: str(sw.r) ?? (sw.s !== undefined ? str(sw.s) : undefined), repeat: sw.c === true };
		case 'MERGEFIELD':
			return { ...fields, kind: 'mergefield', name: arg0 ?? '' };
		case 'INCLUDEPICTURE':
			return { ...fields, kind: 'includepicture', path: arg0 ?? '', linked: sw.d === true };
		case 'PAGE': case 'NUMPAGES':
			return { ...fields, kind: parsed.name === 'PAGE' ? 'page' : 'numpages', format: str(sw['*']) };
		default:
			return { ...fields, kind: 'other', format: str(sw['*']) };
	}
}
