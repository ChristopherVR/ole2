import { readFileSync } from 'node:fs';
import { describe, expect, it } from 'vitest';
import { parseDoc } from '../src/index.js';
import type { DocParagraph } from '../src/doc-document.js';
import { UnsupportedOle2EditError } from '../src/ole2-document-base.js';
import { unwrapDocBytes } from '../src/ole-document-doc-cfb.js';
import { readDocFib } from '../src/ole-document-doc-fib.js';
import { parseBteTable, buildBteTableBytes } from '../src/ole-document-doc-fkp.js';
import { buildClxBytes, parseDocClx } from '../src/ole-document-doc-pieces.js';
import { readDocParagraphAlignments, writeDocParagraphAlignment } from '../src/ole-document-doc-paragraph-format.js';

const fixture = () => new Uint8Array(readFileSync(new URL('./fixtures/doc/paragraph-alignment.doc', import.meta.url)));
function layout(input = fixture()) {
	const cfb = unwrapDocBytes(input)!, word = cfb.wordDocBytes.slice(), fib = readDocFib(word);
	const papx = parseBteTable(cfb.tableBytes, fib.plcfbtePapx), page = papx.pns[0]! * 512;
	const count = word[page + 511]!, header = page + (count + 1) * 4;
	const slots = Array.from({ length: count }, (_, i) => header + i * 13);
	const slot = slots.find((s) => {
		const blob = page + word[s]! * 2, length = word[blob] === 0 ? word[blob + 1]! * 2 + 2 : word[blob]! * 2;
		return word.subarray(blob, blob + length).some((v, i, bytes) => v === 0x61 && bytes[i + 1] === 0x24);
	})!;
	const blob = page + word[slot]! * 2, prefix = word[blob] === 0 ? 2 : 1;
	return { cfb, word, fib, papx, page, slots, slot, blob, prls: blob + prefix + 2 };
}
function altered(variant: string) {
	const { cfb, word, slot, slots, prls, fib, page } = layout();
	if (variant === 'duplicate') word.set([0x61, 0x24], prls);
	if (variant === 'legacy-only') word.set([5, 0x24], prls + 3);
	if (variant === 'mismatch') word[prls + 2] = 0;
	if (variant === 'reverse') { word.set([0x61, 0x24], prls); word.set([3, 0x24], prls + 3); }
	if (variant === 'invalid-old') word[prls + 5] = 10;
	if (variant === 'opaque') word[prls + 6] = 0x99; // Unknown same-width revision SPRM.
	if (variant === 'shared') word[slots.find((s) => s !== slot)!] = word[slot]!;
	if (variant === 'chpx-alias') {
		const table = cfb.tableBytes.slice(), bte = parseBteTable(table, fib.plcfbteChpx);
		new DataView(table.buffer).setUint32(fib.plcfbteChpx.fc + (bte.pns.length + 1) * 4, page / 512, true);
		return cfb.rewrap(word, table);
	}
	return cfb.rewrap(word, cfb.tableBytes);
}

describe('direct logical PAPX paragraph alignment', () => {
	it('reads the native exception without inventing inherited defaults and preserves public construction compatibility', () => {
		const compatible: DocParagraph = { index: 0, text: 'caller-created' };
		expect(compatible.directAlignment).toBeUndefined();
		const doc = parseDoc(fixture());
		expect(doc.paragraphs.map((p) => p.directAlignment)).toStrictEqual([undefined, 'center', undefined]);
		expect(doc.capabilities.write).toContain('existing-direct-paragraph-alignment');
		expect(Object.isFrozen(doc.paragraphs[1])).toBe(true);
	});

	it('updates both native compatibility operands, retains all other bytes and restores byte identity with the stable handle', () => {
		const input = fixture(), doc = parseDoc(input), paragraph = doc.paragraphs[1]!;
		const texts = doc.paragraphs.map((p) => p.text), fonts = paragraph.runs.map((r) => r.sprms), style = paragraph.styleIndex;
		paragraph.directAlignment = 'center';
		expect(doc.dirty).toBe(false);
		expect(doc.revision).toBe(0);
		paragraph.directAlignment = 'justify';
		expect(paragraph.directAlignment).toBe('justify');
		expect(parseDoc(doc.serialize()).paragraphs[1]!.directAlignment).toBe('justify');
		expect(doc.serialize().filter((v, i) => v !== input[i]).length).toBe(2);
		expect(paragraph.runs.map((r) => r.sprms)).toStrictEqual(fonts);
		expect(paragraph.styleIndex).toBe(style);
		expect(doc.paragraphs.map((p) => p.text)).toStrictEqual(texts);
		paragraph.directAlignment = 'center';
		expect(doc.serialize()).toStrictEqual(input);
	});

	it('permits logical start/end only on a sole modern slot, without inventing physical bidi resolution', () => {
		const { cfb, word, prls } = layout();
		word.set([5, 0x24], prls); // Synthetic existing keep-together slot in place of legacy mirror.
		const input = cfb.rewrap(word, cfb.tableBytes), doc = parseDoc(input), paragraph = doc.paragraphs[1]!;
		for (const value of ['start', 'end', 'justify', 'center'] as const) {
			paragraph.directAlignment = value;
			expect(parseDoc(doc.serialize()).paragraphs[1]!.directAlignment).toBe(value);
		}
		expect(doc.serialize()).toStrictEqual(input);
	});

	it('refuses direction-dependent mirrored writes and invalid or coercible values without mutation', () => {
		const input = fixture(), doc = parseDoc(input), paragraph = doc.paragraphs[1]!;
		let coercions = 0;
		const coercible = { toString() { coercions++; paragraph.text = 'Not committed'; return 'justify'; } };
		for (const value of ['start', 'end', 'left', 'right', '', undefined, null, 3, coercible]) {
			expect(() => { paragraph.directAlignment = value as 'center'; }).toThrow(UnsupportedOle2EditError);
			expect(doc.serialize()).toStrictEqual(input);
			expect(doc.dirty).toBe(false);
		}
		expect(coercions).toBe(0);
		expect(() => { paragraph.directAlignment = 'end'; }).toThrow('direction-dependent-alignment');
	});

	it('refuses missing, duplicate, legacy-only, mismatched, reversed, invalid, opaque and shared formatting', () => {
		for (const variant of ['duplicate', 'legacy-only', 'mismatch', 'reverse', 'invalid-old', 'opaque', 'shared', 'chpx-alias']) {
			const input = altered(variant), doc = parseDoc(input);
			expect(() => { doc.paragraphs[1]!.directAlignment = 'justify'; }).toThrow(UnsupportedOle2EditError);
			expect(doc.serialize()).toStrictEqual(input);
			expect(doc.dirty).toBe(false);
		}
		const doc = parseDoc(fixture());
		expect(() => { doc.paragraphs[0]!.directAlignment = 'center'; }).toThrow('inherited-formatting');
	});

	it('retains current paragraph ownership after a text or reentrant RHS edit and invalidates run snapshots', () => {
		const doc = parseDoc(fixture()), paragraph = doc.paragraphs[1]!, oldRun = paragraph.runs[0]!;
		const value = { get alignment() { doc.paragraphs[0]!.text = 'Other plain paragraph.'; return 'justify' as const; } };
		paragraph.directAlignment = value.alignment;
		expect(doc.paragraphs[0]!.text).toBe('Other plain paragraph.');
		expect(paragraph.directAlignment).toBe('justify');
		expect(() => { oldRun.directBold = false; }).toThrow('stale-run');
		paragraph.directAlignment = 'center';
		expect(paragraph.directAlignment).toBe('center');
	});

	it('rejects alias pieces in another story, PRM overlays and oversized formatting tables', () => {
		for (const variant of ['alias', 'prm', 'budget']) {
			const { cfb, word, fib } = layout(), clx = parseDocClx(cfb.tableBytes, fib.clx), pieces = clx.pieces;
			if (variant === 'prm') pieces[0]!.prm = 2; // Valid Prm0 overlay, not a missing Prc reference.
			if (variant === 'alias') {
				const end = pieces.at(-1)!.cpEnd;
				pieces.push({ ...pieces[0]!, cpStart: end, cpEnd: end + 1, fc: pieces[0]!.fc + 23 });
				new DataView(word.buffer).setInt32(fib.ccpTextOffset + 4, 1, true);
			}
			const append = buildClxBytes(pieces, clx.prcBytes), table = new Uint8Array(cfb.tableBytes.length + append.length);
			table.set(cfb.tableBytes); table.set(append, cfb.tableBytes.length);
			const view = new DataView(word.buffer);
			view.setUint32(fib.fibRgFcLcbOffset + 33 * 8, cfb.tableBytes.length, true);
			view.setUint32(fib.fibRgFcLcbOffset + 33 * 8 + 4, append.length, true);
			if (variant === 'budget') view.setUint32(fib.fibRgFcLcbOffset + 13 * 8 + 4, 65_537 * 8 + 4, true);
			const input = cfb.rewrap(word, table), doc = parseDoc(input);
			expect(() => { doc.paragraphs[1]!.directAlignment = 'justify'; }).toThrow(UnsupportedOle2EditError);
			expect(doc.serialize()).toStrictEqual(input);
			expect(doc.dirty).toBe(false);
		}
		expect(() => writeDocParagraphAlignment(new Uint8Array(), -1, 'center')).toThrow('invalid-formatting');
		expect(() => readDocParagraphAlignments(new Uint8Array())).toThrow();
	});

	it('refuses section-exception aliases and malformed PAPX blobs before changing bytes', () => {
		for (const variant of ['section-alias', 'zero-size', 'header-offset', 'inverted-range']) {
			const { cfb, word, fib, page, slot, blob } = layout(), table = cfb.tableBytes.slice();
			if (variant === 'section-alias') {
				const n = (fib.sed.lcb - 4) / 16;
				new DataView(table.buffer).setInt32(fib.sed.fc + (n + 1) * 4 + 2, blob, true);
			}
			if (variant === 'zero-size') { word[blob] = 0; word[blob + 1] = 0; }
			if (variant === 'header-offset') word[slot] = 1;
			if (variant === 'inverted-range') new DataView(word.buffer).setInt32(page + 4, 0, true);
			const input = cfb.rewrap(word, table), doc = parseDoc(input);
			expect(() => { doc.paragraphs[1]!.directAlignment = 'justify'; }).toThrow(UnsupportedOle2EditError);
			expect(doc.serialize()).toStrictEqual(input);
			expect(doc.dirty).toBe(false);
		}
	});

	it('bounds aggregate SPRM allocation even when page and physical-run counts are below their limits', () => {
		const { cfb, fib, papx } = layout(), pageStart = Math.ceil(cfb.wordDocBytes.length / 512), count = 3_000;
		const word = new Uint8Array((pageStart + count) * 512);
		word.set(cfb.wordDocBytes);
		const wordView = new DataView(word.buffer), firstFc = papx.fcs.at(-1)!;
		for (let i = 0; i < count; i++) {
			const page = (pageStart + i) * 512;
			wordView.setInt32(page, firstFc + i, true);
			wordView.setInt32(page + 4, firstFc + i + 1, true);
			word[page + 8] = 11; // One PAPX blob at byte22, past the21-byte header.
			word[page + 22] = 0; word[page + 23] = 151; // istd2 +100three-byte SPRMs.
			for (let j = 0; j < 100; j++) word.set([5, 0x24, 0], page + 26 + j * 3);
			word[page + 511] = 1;
			papx.fcs.push(firstFc + i + 1); papx.pns.push(pageStart + i);
		}
		const bte = buildBteTableBytes(papx), table = new Uint8Array(cfb.tableBytes.length + bte.length);
		table.set(cfb.tableBytes); table.set(bte, cfb.tableBytes.length);
		wordView.setUint32(fib.fibRgFcLcbOffset + 13 * 8, cfb.tableBytes.length, true);
		wordView.setUint32(fib.fibRgFcLcbOffset + 13 * 8 + 4, bte.length, true);
		const input = cfb.rewrap(word, table), doc = parseDoc(input);
		expect(() => readDocParagraphAlignments(input)).toThrow('resource-limit');
		expect(() => { doc.paragraphs[1]!.directAlignment = 'justify'; }).toThrow(UnsupportedOle2EditError);
		expect(doc.serialize()).toStrictEqual(input);
		expect(doc.dirty).toBe(false);
	});
});
