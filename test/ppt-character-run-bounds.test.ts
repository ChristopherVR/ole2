import { readFileSync } from 'node:fs';
import { describe, expect, it } from 'vitest';
import { PptDocument } from '../src/ppt-document.js';
import { readPptCharacterRuns } from '../src/legacy-ppt-text-style.js';
import { readCompoundFileStream, replaceCompoundFileStream } from '../src/ole2-stream-edit.js';
import { resizeCompoundFileStream } from '../src/ole2-stream-resize.js';
import { readRecordOrThrow } from '../src/legacy-ppt-record-stream.js';
import { readPptSlideShapes, pptShapeChildren } from '../src/legacy-ppt-shape-reader.js';
import { buildPersistDirectory } from '../src/ppt/persist-directory.js';
import { RT, OA } from '../src/legacy-ppt-record-types.js';
const load = () => new Uint8Array(readFileSync(new URL('./fixtures/ppt/native-text.ppt', import.meta.url)));
const viewOf = (bytes: Uint8Array) => new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
function atom(type: number, payload: Uint8Array): Uint8Array {
	const bytes = new Uint8Array(8 + payload.length), view = viewOf(bytes); view.setUint16(2, type, true); view.setUint32(4, payload.length, true); bytes.set(payload, 8); return bytes;
}
/** Owned native records with a valid incremental save; no third-party bytes. */
function replaceTitleStyle(makeRecords: (stylePayload: Uint8Array) => Uint8Array): Uint8Array {
	const input = load(), stream = readCompoundFileStream(input, ['PowerPoint Document'])!, user = readCompoundFileStream(input, ['Current User'])!, view = viewOf(stream);
	const oldEditOffset = viewOf(user).getUint32(16, true), chain = buildPersistDirectory(view, oldEditOffset);
	const slide = readRecordOrThrow(view, chain.directory.get(15)!), edit = readRecordOrThrow(view, oldEditOffset);
	const style = readRecordOrThrow(view, readPptCharacterRuns(input)[0]!.runs[0]!.styleHeaderOffset);
	const replacement = makeRecords(stream.subarray(style.dataOffset, style.dataOffset + style.recLen));
	const delta = replacement.length - style.recLen - 8, oldSlide = stream.subarray(slide.headerOffset, slide.dataOffset + slide.recLen), nextSlide = new Uint8Array(oldSlide.length + delta);
	const position = style.headerOffset - slide.headerOffset;
	nextSlide.set(oldSlide.subarray(0, position)); nextSlide.set(replacement, position); nextSlide.set(oldSlide.subarray(position + 8 + style.recLen), position + replacement.length);
	const ancestors = [slide];
	for (let i = 0; i < ancestors.length; i++) for (const child of pptShapeChildren(view, ancestors[i]!)) {
		if ((child.recVer === 15 || child.recType === OA.ClientTextbox) && child.dataOffset <= style.headerOffset && child.dataOffset + child.recLen >= style.dataOffset + style.recLen) ancestors.push(child);
	}
	for (const parent of ancestors) viewOf(nextSlide).setUint32(parent.headerOffset - slide.headerOffset + 4, parent.recLen + delta, true);
	const dirOffset = stream.length + nextSlide.length, editOffset = dirOffset + 16, oldEdit = stream.subarray(edit.headerOffset, edit.dataOffset + edit.recLen), out = new Uint8Array(editOffset + oldEdit.length), ov = viewOf(out);
	out.set(stream); out.set(nextSlide, stream.length); out.set(oldEdit, editOffset);
	ov.setUint16(dirOffset + 2, RT.PersistDirectoryAtom, true); ov.setUint32(dirOffset + 4, 8, true); ov.setUint32(dirOffset + 8, (1 << 20) | 15, true); ov.setUint32(dirOffset + 12, stream.length, true);
	ov.setUint32(editOffset + 16, oldEditOffset, true); ov.setUint32(editOffset + 20, dirOffset, true);
	const nextUser = user.slice(); viewOf(nextUser).setUint32(16, editOffset, true);
	const result = resizeCompoundFileStream(replaceCompoundFileStream(input, ['Current User'], nextUser), ['PowerPoint Document'], out);
	if (!result.ok) throw new Error(result.reason); return result.bytes;
}
describe('PPT character-format bounds and ambiguity', () => {
	it.each([
		['truncated CF exception', (payload: Uint8Array) => atom(RT.StyleTextPropAtom, payload.subarray(0, payload.length - 1)), /Truncated/],
		['trailing style bytes', (payload: Uint8Array) => { const next = new Uint8Array(payload.length + 2); next.set(payload); return atom(RT.StyleTextPropAtom, next); }, /Unconsumed/],
		['duplicate style atoms', (payload: Uint8Array) => { const record = atom(RT.StyleTextPropAtom, payload), next = new Uint8Array(record.length * 2); next.set(record); next.set(record, record.length); return next; }, /ambiguous/],
		['duplicate text atoms', (payload: Uint8Array) => { const text = atom(RT.TextBytesAtom, Uint8Array.of(65)), style = atom(RT.StyleTextPropAtom, payload), next = new Uint8Array(text.length + style.length); next.set(text); next.set(style, text.length); return next; }, /ambiguous/],
	] as const)('preserves valid text and reports unsupported %s', (_label, makeRecords, diagnostic) => {
		const input = replaceTitleStyle(makeRecords), document = new PptDocument(input), title = document.slides[0]!.texts[0]!;
		expect(title.text).toBe('Native title fixture'); expect(title.runsStatus).toBe('unsupported'); expect(title.runsDiagnostic).toMatch(diagnostic);
		expect(title.runs).toEqual([]); expect(document.serialize()).toEqual(input); expect(document.dirty).toBe(false);
	});
	it('reads an implicit-only terminal run but refuses to edit it', () => {
		const input = replaceTitleStyle(payload => {
			const next = new Uint8Array(payload.length + 10); next.set(payload); viewOf(next).setUint32(10, 20, true);
			viewOf(next).setUint32(payload.length, 1, true); viewOf(next).setUint32(payload.length + 4, 0x20000, true); viewOf(next).setInt16(payload.length + 8, 24, true);
			return atom(RT.StyleTextPropAtom, next);
		});
		const document = new PptDocument(input), title = document.slides[0]!.texts[0]!, terminal = title.runs[1]!;
		expect(title.runsStatus).toBe('decoded'); expect([terminal.start, terminal.end, terminal.text]).toEqual([20, 20, '']);
		terminal.directFontSizePoints = 24; expect(document.dirty).toBe(false);
		expect(() => { terminal.directFontSizePoints = 30; }).toThrow(/implicit paragraph mark/);
		expect(document.serialize()).toEqual(input); expect(document.revision).toBe(0);
	});
	it('patches an existing literal multi-run slot while preserving its neighbor and unknown bytes', () => {
		const input = replaceTitleStyle(payload => {
			const next = new Uint8Array(payload.length + 14); next.set(payload.subarray(0, 10)); next.set(payload.subarray(10), 10); next.set(payload.subarray(10), 24);
			viewOf(next).setUint32(10, 3, true); viewOf(next).setInt16(22, 20, true);
			viewOf(next).setUint32(24, 18, true); viewOf(next).setInt16(36, 24, true);
			return atom(RT.StyleTextPropAtom, next);
		});
		const document = new PptDocument(input), runs = document.slides[0]!.texts[0]!.runs;
		expect(runs.map(run => [run.start, run.end, run.text, run.directFontSizePoints])).toEqual([[0, 3, 'Nat', 20], [3, 20, 'ive title fixture', 24]]);
		runs[1]!.directFontSizePoints = 30;
		expect(runs[0]!.directFontSizePoints).toBe(20); expect(runs[1]!.directFontSizePoints).toBe(30);
		expect(document.serialize().reduce((count, byte, index) => count + Number(byte !== input[index]), 0)).toBe(1);
	});
	it('refuses literal size edits in mirrored text-box context', () => {
		const input = load(), stream = readCompoundFileStream(input, ['PowerPoint Document'])!.slice(), view = viewOf(stream);
		const shape = readPptSlideShapes(input)[0]!.shapes[0]!, opt = pptShapeChildren(view, readRecordOrThrow(view, shape.headerOffset)).find(record => record.recType === OA.FOPT)!;
		// A hostile mirror declaration remains byte-preserved and blocks writes.
		view.setUint16(opt.dataOffset, (view.getUint16(opt.dataOffset, true) & 0xc000) | 0x3a9, true);
		const mirrored = replaceCompoundFileStream(input, ['PowerPoint Document'], stream), document = new PptDocument(mirrored), run = document.slides[0]!.texts[0]!.runs[0]!;
		run.directFontSizePoints = 28; expect(document.dirty).toBe(false);
		expect(() => { run.directFontSizePoints = 32; }).toThrow(/mirror/i); expect(document.serialize()).toEqual(mirrored);
	});
});
