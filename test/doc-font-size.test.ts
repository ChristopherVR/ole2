import { readFileSync } from 'node:fs';
import { describe, expect, it } from 'vitest';
import { DocDocument } from '../src/doc-document.js';
import { parseDoc } from '../src/index.js';
import { UnsupportedOle2EditError } from '../src/ole2-document-base.js';
import { unwrapDocBytes } from '../src/ole-document-doc-cfb.js';
import { readDocFib } from '../src/ole-document-doc-fib.js';
import { parseBteTable } from '../src/ole-document-doc-fkp.js';
import { writeDocCharacterRunFontSize } from '../src/ole-document-doc-runs.js';

const original = () => new Uint8Array(readFileSync(new URL('./fixtures/doc/rich-runs.doc', import.meta.url)));
// Synthetic fixture only: replace a six-byte revision ID SPRM with a four-byte
// size SPRM. Native validation uses a separately generated Word size fixture.
function prepared(variant = 'valid') {
	const cfb = unwrapDocBytes(original())!, word = cfb.wordDocBytes.slice(), fib = readDocFib(word);
	const bte = parseBteTable(cfb.tableBytes, fib.plcfbteChpx), page = bte.pns[0]! * 512;
	const count = word[page + 511]!, rgb = page + (count + 1) * 4;
	const slots = Array.from({ length: count }, (_, i) => rgb + i);
	const slot = slots.find((s) => {
		const blob = page + word[s]! * 2;
		return word[blob + word[blob]! - 1] === 0x35 && word[blob + word[blob]!] === 0x08;
	}) ?? slots.find((s) => {
		const blob = page + word[s]! * 2;
		return word.subarray(blob + 1, blob + 1 + word[blob]!).some((v, i, bytes) => v === 0x35 && bytes[i + 1] === 8);
	})!;
	const blob = page + word[slot]! * 2, length = word[blob]!;
	word.copyWithin(blob + 5, blob + 7, blob + 1 + length);
	word[blob] = length - 2;
	word.set([0x43, 0x4a, 36, 0], blob + 1);
	if (variant === 'invalid-old') word[blob + 3] = 1;
	if (variant === 'opaque') word[blob + 5] = 0x99;
	if (variant === 'duplicate') word.set([0x43, 0x4a, 24, 0, 0x3e, 0x2a], blob + 5);
	if (variant === 'shared') word[slots.find((s) => word[s] !== word[slot])!] = word[slot]!;
	return cfb.rewrap(word, cfb.tableBytes);
}

describe('exclusive direct DOC font size', () => {
	it('edits the owned native Word size fixture through the primary model API and restores all bytes', () => {
		const input = new Uint8Array(readFileSync(new URL('./fixtures/doc/rich-size-runs.doc', import.meta.url)));
		const doc = parseDoc(input), paragraph = doc.paragraphs[1]!;
		expect(paragraph.runs[0]!.directFontSizePoints).toBe(18);
		for (const size of [12, 13.5, 130, 1, 1638, 18]) {
			paragraph.runs[0]!.directFontSizePoints = size;
			expect(parseDoc(doc.serialize()).paragraphs[1]!.runs[0]!.directFontSizePoints).toBe(size);
		}
		expect(doc.serialize()).toStrictEqual(input);
	});

	it('changes only the two-byte operand, refreshes raw SPRMs and preserves text and all other bytes', () => {
		const input = prepared(), doc = new DocDocument(input), paragraph = doc.paragraphs[1]!, run = paragraph.runs[0]!;
		expect(run.text).toBe('Bold text');
		expect(run.directFontSizePoints).toBe(18);
		const before = run.sprms;
		run.directFontSizePoints = 130;
		expect(run.directFontSizePoints).toBe(130);
		expect(run.sprms.find((p) => p.opcode === 0x4a43)!.operand).toStrictEqual([4, 1]);
		expect(run.sprms.filter((p) => p.opcode !== 0x4a43)).toStrictEqual(before.filter((p) => p.opcode !== 0x4a43));
		const output = doc.serialize();
		expect(output.length).toBe(input.length);
		expect(output.filter((v, i) => v !== input[i]).length).toBe(2);
		expect(new DocDocument(output).paragraphs[1]!.runs[0]!.directFontSizePoints).toBe(130);
		paragraph.runs[0]!.directFontSizePoints = 13.5;
		expect(paragraph.runs[0]!.directFontSizePoints).toBe(13.5);
		expect(doc.paragraphs.map((p) => p.text)).toStrictEqual(new DocDocument(input).paragraphs.map((p) => p.text));
	});

	it('accepts exact half-point endpoints and leaves no-op documents clean and handles usable', () => {
		const input = prepared(), doc = new DocDocument(input), run = doc.paragraphs[1]!.runs[0]!;
		run.directFontSizePoints = 18;
		expect(doc.dirty).toBe(false);
		expect(doc.revision).toBe(0);
		expect(doc.serialize()).toStrictEqual(input);
		for (const points of [1, 1.5, 1638]) {
			doc.paragraphs[1]!.runs[0]!.directFontSizePoints = points;
			expect(new DocDocument(doc.serialize()).paragraphs[1]!.runs[0]!.directFontSizePoints).toBe(points);
		}
	});

	it('rejects invalid primitives and objects without invoking coercion or mutating state', () => {
		const input = prepared(), doc = new DocDocument(input), run = doc.paragraphs[1]!.runs[0]!;
		let coercions = 0;
		const object = { valueOf() { coercions++; doc.paragraphs[0]!.text = 'Other plain paragraph.'; return 13.5; } };
		for (const value of [undefined, NaN, Infinity, -Infinity, 0.5, 1638.5, 13.25, '13.5', null, object]) {
			expect(() => { run.directFontSizePoints = value as number; }).toThrow(UnsupportedOle2EditError);
			try { run.directFontSizePoints = value as number; }
			catch (error) { expect(error).toMatchObject({ name: 'UnsupportedOle2EditError', reason: 'invalid-formatting' }); }
			expect(doc.serialize()).toStrictEqual(input);
			expect(doc.dirty).toBe(false);
		}
		expect(coercions).toBe(0);
	});

	it('refuses stale handles after a size, flag or reentrant RHS edit', () => {
		const doc = new DocDocument(prepared()), run = doc.paragraphs[1]!.runs[0]!;
		run.directFontSizePoints = 13.5;
		expect(() => { run.directFontSizePoints = 18; }).toThrow('stale-run');
		const afterSize = doc.paragraphs[1]!.runs[0]!;
		afterSize.directBold = false;
		expect(() => { afterSize.directFontSizePoints = 18; }).toThrow('stale-run');
		const current = doc.paragraphs[1]!.runs[0]!;
		const reentrant = { get points() { doc.paragraphs[0]!.text = 'Other plain paragraph.'; return 18; } };
		expect(() => { current.directFontSizePoints = reentrant.points; }).toThrow('stale-run');
		expect(doc.paragraphs[0]!.text).toBe('Other plain paragraph.');
		expect(doc.paragraphs[1]!.runs[0]!.directFontSizePoints).toBe(13.5);
	});

	it('refuses inherited, malformed, duplicate, opaque and shared operands transactionally', () => {
		for (const input of [original(), prepared('invalid-old'), prepared('duplicate'), prepared('opaque'), prepared('shared')]) {
			const doc = new DocDocument(input);
			expect(() => { doc.paragraphs[1]!.runs[0]!.directFontSizePoints = 13.5; }).toThrow(UnsupportedOle2EditError);
			expect(doc.serialize()).toStrictEqual(input);
			expect(doc.dirty).toBe(false);
		}
	});

	it('bounds direct helper inputs before parsing and refuses partial logical runs', () => {
		const input = prepared(), run = new DocDocument(input).paragraphs[1]!.runs[0]!;
		for (const value of [NaN, Infinity, 13.25, 0, 1639]) expect(() => writeDocCharacterRunFontSize(new Uint8Array(), run.cpStart, run.cpEnd, value)).toThrow('invalid-formatting');
		expect(() => writeDocCharacterRunFontSize(input, run.cpStart + 1, run.cpEnd, 13.5)).toThrow('unsupported-formatting');
		expect(() => writeDocCharacterRunFontSize(input, NaN, run.cpEnd, 13.5)).toThrow('invalid-formatting');
	});
});
