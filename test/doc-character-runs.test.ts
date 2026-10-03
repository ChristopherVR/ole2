import { readFileSync } from 'node:fs';
import { describe, expect, it } from 'vitest';
import { DocDocument } from '../src/doc-document.js';
import type { DocParagraph } from '../src/doc-document.js';
import { UnsupportedOle2EditError } from '../src/ole2-document-base.js';
import { unwrapDocBytes } from '../src/ole-document-doc-cfb.js';
import { readDocFib } from '../src/ole-document-doc-fib.js';
import { parseBteTable } from '../src/ole-document-doc-fkp.js';
import { parseDocClx, buildClxBytes } from '../src/ole-document-doc-pieces.js';
import { readDocCharacterRuns, writeDocCharacterRunFlag } from '../src/ole-document-doc-runs.js';

const fixture = () => new Uint8Array(readFileSync(new URL('./fixtures/doc/rich-runs.doc', import.meta.url)));
function layout(input = fixture()) {
	const cfb = unwrapDocBytes(input)!, word = cfb.wordDocBytes.slice(), fib = readDocFib(word);
	const bte = parseBteTable(cfb.tableBytes, fib.plcfbteChpx), page = bte.pns[0]! * 512;
	const count = word[page + 511]!, rgb = page + (count + 1) * 4;
	const offsets = Array.from({ length: count }, (_, i) => ({ slot: rgb + i, blob: page + word[rgb + i]! * 2 }));
	const target = offsets.find((r) => {
		for (let at = r.blob + 1; at + 2 < r.blob + 1 + word[r.blob]!; at++)
			if (word[at] === 0x35 && word[at + 1] === 0x08) return true;
		return false;
	})!;
	return { cfb, word, fib, offsets, target };
}

describe('DOC native character runs and exclusive formatting edits', () => {
	it('maps native mixed formatting, preserves opaque records and does not claim resolved styles', () => {
		const oldPublicShape: DocParagraph = { index: 0, text: 'Compatible' };
		expect(oldPublicShape.text).toBe('Compatible');
		const doc = new DocDocument(fixture()), paragraph = doc.paragraphs[1]!;
		expect(paragraph.runs.map((r) => r.text)).toStrictEqual(['Bold text', ', ', 'italic text', ', plain text.']);
		expect(paragraph.runs.map((r) => r.text).join('')).toBe(paragraph.text);
		expect(paragraph.styleIndex).toBe(0);
		const bold = paragraph.runs[0]!;
		expect(bold.directBold).toBeUndefined(); // Native operand129 depends on style.
		expect(bold.sprms.find((p) => p.opcode === 0x0835)!.operand).toStrictEqual([129]);
		expect(Object.isFrozen(bold.sprms[0]!.operand)).toBe(true);
		expect(paragraph.runs.at(-1)!.directItalic).toBeUndefined(); // No explicit operand.
	});

	it('forces absolute bold/italic with exactly one changed container byte and reloads after multiple edits', () => {
		const input = fixture(), doc = new DocDocument(input), paragraph = doc.paragraphs[1]!;
		const before = paragraph.runs[0]!.sprms, bold = paragraph.runs[0]!;
		bold.directBold = false;
		expect(bold.directBold).toBe(false);
		expect(bold.sprms.find((p) => p.opcode === 0x0835)!.operand).toStrictEqual([0]);
		expect(bold.sprms.filter((p) => p.opcode !== 0x0835)).toStrictEqual(before.filter((p) => p.opcode !== 0x0835));
		const edited = doc.serialize();
		expect(edited.length).toBe(input.length);
		expect(edited.filter((value, i) => value !== input[i]).length).toBe(1);
		expect(() => { bold.directBold = true; }).toThrow('stale-run');
		const revision = doc.revision;
		paragraph.runs[0]!.directBold = false;
		expect(doc.revision).toBe(revision);
		paragraph.runs[2]!.directItalic = false;
		paragraph.runs[0]!.directBold = true;
		const reloaded = new DocDocument(doc.serialize());
		expect(reloaded.paragraphs[1]!.runs[0]!.directBold).toBe(true);
		expect(reloaded.paragraphs[1]!.runs[2]!.directItalic).toBe(false);
		expect(reloaded.paragraphs.map((p) => p.text)).toStrictEqual(doc.paragraphs.map((p) => p.text));
		doc.paragraphs[0]!.text = 'Other plain paragraph.'; // Same length preserves rich runs.
		expect(paragraph.runs[0]!.directBold).toBe(true);
	});

	it('refuses missing slots, shared blobs, opaque semantics and duplicate operands transactionally', () => {
		const cases: Uint8Array[] = [];
		for (const variant of ['opaque', 'duplicate', 'shared', 'invalid-toggle']) {
			const { cfb, word, target, offsets } = layout();
			if (variant === 'opaque') word[target.blob + 1] = 0x99; // Unknown same-size SPRM preserves framing.
			if (variant === 'duplicate') word.set([0x35, 8, 1, 0x35, 8, 1], target.blob + 1);
			if (variant === 'shared') word[offsets.find((r) => r.blob !== target.blob)!.slot] = word[target.slot]!;
			if (variant === 'invalid-toggle') {
				for (let at = target.blob + 1; at < target.blob + word[target.blob]!; at++)
					if (word[at] === 0x35 && word[at + 1] === 8) word[at + 2] = 2;
			}
			cases.push(cfb.rewrap(word, cfb.tableBytes));
		}
		for (const input of cases) {
			const doc = new DocDocument(input);
			expect(() => { doc.paragraphs[1]!.runs.find((r) => r.text === 'Bold text')!.directBold = false; }).toThrow(UnsupportedOle2EditError);
			expect(doc.serialize()).toStrictEqual(input);
			expect(doc.dirty).toBe(false);
		}
		const doc = new DocDocument(fixture()), plain = doc.paragraphs[1]!.runs[1]!;
		expect(() => { plain.directBold = true; }).toThrow(UnsupportedOle2EditError);
		expect(doc.dirty).toBe(false);
	});

	it('refuses alias pieces in a non-main story and malformed FKP ranges', () => {
		const { cfb, word, fib } = layout(), clx = parseDocClx(cfb.tableBytes, fib.clx);
		const run = readDocCharacterRuns(fixture()).find((r) => r.text === 'Bold text')!;
		const piece = clx.pieces.find((p) => p.cpStart <= run.cpStart && p.cpEnd >= run.cpEnd)!;
		const end = clx.pieces.at(-1)!.cpEnd;
		clx.pieces.push({ ...piece, fc: piece.fc + (run.cpStart - piece.cpStart) * (piece.compressed ? 1 : 2), cpStart: end, cpEnd: end + 9 });
		const appended = buildClxBytes(clx.pieces, clx.prcBytes), table = new Uint8Array(cfb.tableBytes.length + appended.length);
		table.set(cfb.tableBytes); table.set(appended, cfb.tableBytes.length);
		const view = new DataView(word.buffer);
		view.setUint32(fib.fibRgFcLcbOffset + 33 * 8, cfb.tableBytes.length, true);
		view.setUint32(fib.fibRgFcLcbOffset + 33 * 8 + 4, appended.length, true);
		view.setInt32(fib.ccpTextOffset + 4, 9, true);
		const alias = cfb.rewrap(word, table);
		expect(() => writeDocCharacterRunFlag(alias, run.cpStart, run.cpEnd, 'bold', false)).toThrow('shared-formatting');
		const malformed = layout();
		const bte = parseBteTable(malformed.cfb.tableBytes, malformed.fib.plcfbteChpx);
		new DataView(malformed.word.buffer).setInt32(bte.pns[0]! * 512 + 4, 0, true);
		expect(() => readDocCharacterRuns(malformed.cfb.rewrap(malformed.word, malformed.cfb.tableBytes))).toThrow();
	});
});
