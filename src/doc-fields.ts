/** Full PlcfFld parsing for every story: begin (0x13), separator (0x14) and end
 * (0x15) marks nested into a field tree with typed instructions
 * ([MS-DOC] 2.8.25 PlcFld, 2.9.88 Fld, 2.9.89 Fldch). The main-story shifter in
 * `ole-document-doc-fields.ts` edits the same table; this module only reads. */
import { DocResourceLimitError } from './ole-document-doc-pieces.js';
import { FCLCB, tableRange } from './doc-objects-context.js';
import type { DocObjectsContext } from './doc-objects-context.js';
import { parseFieldInstruction, typeField } from './doc-field-instruction.js';
import type { DocField, DocStoryKind } from './doc-objects-types.js';

const FIELD_TABLES: readonly [number, DocStoryKind][] = [
	[FCLCB.fldMom, 'main'], [FCLCB.fldHdr, 'headers'], [FCLCB.fldFtn, 'footnotes'], [FCLCB.fldAtn, 'comments'],
	[FCLCB.fldEdn, 'endnotes'], [FCLCB.fldTxbx, 'textboxes'], [FCLCB.fldHdrTxbx, 'headerTextboxes'],
];

const MAX_DEPTH = 64;

interface Partial {
	begin: number;
	flt: number;
	separator: number | undefined;
	end: number;
	endFlags: number;
	parent: number;
	depth: number;
}

/** Text in `[from, to)` with control marks removed and nested fields replaced by their result. */
function flatten(ctx: DocObjectsContext, from: number, to: number, byBegin: ReadonlyMap<number, Partial>, depth = 0): string {
	let out = '', cp = from;
	for (let i = ctx.text.indexOf('', from); i >= 0 && i < to && depth < MAX_DEPTH; i = ctx.text.indexOf('', i + 1)) {
		const child = byBegin.get(i);
		if (!child || child.end < 0 || child.end >= to) continue;
		out += plain(ctx.text.slice(cp, i));
		if (child.separator !== undefined) out += flatten(ctx, child.separator + 1, child.end, byBegin, depth + 1);
		cp = i = child.end;
		cp++;
	}
	return out + plain(ctx.text.slice(cp, to));
}

function plain(s: string): string {
	return s.replace(/[\u0001\u0008\u0013\u0014\u0015]/gu, '');
}

function buildStory(ctx: DocObjectsContext, bytes: Uint8Array, story: DocStoryKind, base: number, into: Partial[]): void {
	if (bytes.length < 4 || (bytes.length - 4) % 6 !== 0) {
		ctx.warnings.push(`Field table for the ${story} story has an invalid length`);
		return;
	}
	const n = (bytes.length - 4) / 6;
	if (n + into.length > ctx.limits.maxRecords) throw new DocResourceLimitError('Field count exceeds the processing limit');
	const v = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
	const stack: Partial[] = [];
	const stackIndex: number[] = [];
	let previous = -1;
	const first = into.length;
	for (let i = 0; i < n; i++) {
		const cp = base + v.getInt32(i * 4, true);
		const ch = bytes[(n + 1) * 4 + i * 2]! & 0x1f, second = bytes[(n + 1) * 4 + i * 2 + 1]!;
		if (cp <= previous || cp < 0 || cp >= ctx.text.length || ctx.text.charCodeAt(cp) !== ch) {
			ctx.warnings.push(`Field mark ${i} of the ${story} story does not match the text`);
			into.length = first;
			return;
		}
		previous = cp;
		if (ch === 0x13) {
			if (stack.length >= MAX_DEPTH) {
				ctx.warnings.push(`Field nesting in the ${story} story exceeds ${MAX_DEPTH}`);
				into.length = first;
				return;
			}
			const p: Partial = { begin: cp, flt: second, separator: undefined, end: -1, endFlags: 0, parent: stackIndex.at(-1) ?? -1, depth: stack.length };
			stack.push(p); stackIndex.push(into.length); into.push(p);
		} else if (ch === 0x14) {
			const top = stack.at(-1);
			if (top && top.separator === undefined) top.separator = cp;
		} else if (ch === 0x15) {
			const top = stack.pop();
			stackIndex.pop();
			if (top) { top.end = cp; top.endFlags = second; }
		}
	}
	if (stack.length) ctx.warnings.push(`Field table for the ${story} story has unterminated fields`);
}

/** Every complete field in all stories, in begin order; `parent` indexes the returned array. */
export function readDocFields(ctx: DocObjectsContext): DocField[] {
	const partials: Partial[] = [];
	const stories: DocStoryKind[] = [];
	for (const [index, story] of FIELD_TABLES) {
		const bytes = tableRange(ctx, ctx.pair(index));
		const base = ctx.stories.find((s) => s.story === story)?.cpStart;
		if (!bytes || base === undefined) continue;
		const before = partials.length;
		buildStory(ctx, bytes, story, base, partials);
		for (let i = before; i < partials.length; i++) stories[i] = story;
	}
	const kept = partials.map((p, i) => ({ p, i })).filter(({ p }) => p.end >= 0);
	const newIndex = new Map(kept.map(({ i }, k) => [i, k]));
	const byBegin = new Map(partials.map((p) => [p.begin, p]));
	return kept.map(({ p, i }) => {
		const codeEnd = p.separator ?? p.end;
		const instruction = flatten(ctx, p.begin + 1, codeEnd, byBegin).trim();
		const resultRange = p.separator === undefined ? undefined : { cpStart: p.separator + 1, cpEnd: p.end };
		const result = resultRange ? flatten(ctx, resultRange.cpStart, resultRange.cpEnd, byBegin) : '';
		const parsed = parseFieldInstruction(instruction);
		return typeField({
			story: stories[i]!, cpBegin: p.begin, cpSeparator: p.separator, cpEnd: p.end, flt: p.flt,
			parent: p.parent < 0 ? -1 : newIndex.get(p.parent) ?? -1, depth: p.depth,
			instruction, result, resultRange, switches: {}, endFlags: p.endFlags,
		}, parsed);
	});
}
