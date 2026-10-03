import { readFileSync } from 'node:fs';
import { describe, expect, it } from 'vitest';
import { PptDocument, PptText, type PptCharacterRun } from '../src/ppt-document.js';
import { readPptCharacterRuns, editPptCharacterFontSize } from '../src/legacy-ppt-text-style.js';
import { readCompoundFileStream, replaceCompoundFileStream } from '../src/ole2-stream-edit.js';
import { buildPptFile } from '../src/legacy-ppt-writer.js';
import { readPptSlideShapes, pptShapeChildren } from '../src/legacy-ppt-shape-reader.js';
import { readRecordOrThrow } from '../src/legacy-ppt-record-stream.js';
import { OA } from '../src/legacy-ppt-record-types.js';

const load = () => new Uint8Array(readFileSync(new URL('./fixtures/ppt/native-text.ppt', import.meta.url)));
const viewOf = (bytes: Uint8Array) => new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
function mutateStyle(change: (view: DataView, run: ReturnType<typeof readPptCharacterRuns>[number]['runs'][number]) => void): Uint8Array {
	const input = load(), stream = readCompoundFileStream(input, ['PowerPoint Document'])!.slice();
	change(viewOf(stream), readPptCharacterRuns(input)[0]!.runs[0]!);
	return replaceCompoundFileStream(input, ['PowerPoint Document'], stream);
}
describe('PPT direct character-format runs', () => {
	it('reads native literal sizes, unresolved inheritance and exact UTF-16 spans', () => {
		const document = new PptDocument(load()), title = document.slides[0]!.texts[0]!;
		expect(title.runsStatus).toBe('decoded'); expect(title.runsDiagnostic).toBeUndefined();
		expect(title.runs.map(run => [run.start, run.end, run.text, run.directFontSizePoints])).toEqual([[0, 20, title.text, 28]]);
		expect(document.slides[1]!.texts[0]!.runs[0]!.directFontSizePoints).toBe(24);
		const wide = document.slides[1]!.texts[1]!;
		expect(wide.runs.map(run => [run.start, run.end, run.text])).toEqual([[0, 4, '日本語 '], [4, 11, 'café 😀']]);
		expect(document.slides[0]!.texts[1]!.runs[0]!.directFontSizePoints).toBeUndefined();
		expect(document.dirty).toBe(false); expect(document.serialize()).toEqual(load());
	});
	it('retains immutable handles and refreshes getters through font and text edits', () => {
		const input = load(), before = input.slice(), document = new PptDocument(input), title = document.slides[0]!.texts[0]!, run: PptCharacterRun = title.runs[0]!;
		expect(Object.isFrozen(title.runs)).toBe(true); expect(Object.isFrozen(run)).toBe(true);
		expect(Reflect.set(run, 'start', 99)).toBe(false);
		run.directFontSizePoints = 28; expect(document.dirty).toBe(false);
		run.directFontSizePoints = 32; expect(run.directFontSizePoints).toBe(32); expect(document.revision).toBe(1);
		const output = document.serialize(); expect(input).toEqual(before);
		expect(output.reduce((count, byte, index) => count + Number(byte !== input[index]), 0)).toBe(1);
		expect(new PptDocument(output).slides[0]!.texts[0]!.runs[0]!.directFontSizePoints).toBe(32);
		title.text = 'Native title updated'; expect(run.text).toBe(title.text); expect(run.directFontSizePoints).toBe(32);
		run.directFontSizePoints = 28; title.text = 'Native title fixture'; expect(document.serialize()).toEqual(input);
	});
	it.each([undefined, 0, -1, 4001, 1.5, NaN, Infinity])('refuses invalid/removal size %s without changing a previous successful edit', value => {
		const document = new PptDocument(load()), run = document.slides[0]!.texts[0]!.runs[0]!;
		run.directFontSizePoints = 32; const before = document.serialize(), revision = document.revision;
		expect(() => { run.directFontSizePoints = value; }).toThrow(/integer from 1 to 4000/);
		expect(run.directFontSizePoints).toBe(32); expect(document.revision).toBe(revision); expect(document.serialize()).toEqual(before);
	});
	it('refuses inserting inherited font sizes, allowing exact inherited no-ops', () => {
		const document = new PptDocument(load()), run = document.slides[0]!.texts[1]!.runs[0]!;
		run.directFontSizePoints = undefined; expect(document.dirty).toBe(false);
		expect(() => { run.directFontSizePoints = 24; }).toThrow(/inherited/);
		expect(document.serialize()).toEqual(load()); expect(document.revision).toBe(0);
	});
	it.each([
		['zero paragraph count', (v: DataView, r: { styleHeaderOffset: number }) => v.setUint32(r.styleHeaderOffset + 8, 0, true)],
		['excess character count', (v: DataView, r: { runOffset: number }) => v.setUint32(r.runOffset, 0xffffffff, true)],
		['extension mask', (v: DataView, r: { runOffset: number }) => v.setUint32(r.runOffset + 4, 0x00530000, true)],
		['zero literal size', (v: DataView, r: { fontSizeOffset?: number }) => v.setInt16(r.fontSizeOffset!, 0, true)],
	] as const)('keeps valid slide text readable while reporting unsupported %s formatting', (_label, change) => {
		const input = mutateStyle(change), document = new PptDocument(input), text = document.slides[0]!.texts[0]!;
		expect(text.text).toBe('Native title fixture'); expect(text.runsStatus).toBe('unsupported');
		expect(text.runsDiagnostic).toBeTruthy(); expect(text.runs).toEqual([]); expect(document.serialize()).toEqual(input);
	});
	it('validates stale identities and snapshots low-level caller fields exactly once', () => {
		const input = load(), text = readPptCharacterRuns(input)[0]!, run = text.runs[0]!;
		const identity = { slideId: text.slideId, persistId: text.persistId, headerOffset: text.headerOffset,
			runOffset: run.runOffset, styleHeaderOffset: run.styleHeaderOffset, start: run.start, end: run.end, expectedFontSizePoints: 28, fontSizePoints: 32 };
		expect(editPptCharacterFontSize(input, { ...identity, runOffset: run.runOffset + 2 }).status).toBe('unsupported');
		expect(editPptCharacterFontSize(input, { ...identity, expectedFontSizePoints: 29 }).status).toBe('unsupported');
		let reads = 0;
		const edited = editPptCharacterFontSize(input, { ...identity, get fontSizePoints() { return reads++ === 0 ? 32 : -1; } });
		expect(reads).toBe(1); expect(edited.status).toBe('edited'); expect(readPptCharacterRuns(edited.bytes)[0]!.runs[0]!.directFontSizePoints).toBe(32);
	});
	it('reads first-party multi-run literal slots but refuses an undecoded large-anchor context', async () => {
		const generated = await buildPptFile({ widthEmu: 9144000, heightEmu: 5143500, pictures: [], slides: [{ shapes: [{
			kind: 'shape', spt: 202, isConnector: false, anchor: { x: 0, y: 0, w: 3000000, h: 1000000 },
			text: { textType: 4, paragraphs: [{ indentLevel: 0, runs: [{ text: 'A😀', sizePt: 20, bold: true }, { text: 'βZ', sizePt: 24, italic: true }] }] },
		}] }] });
		// First-party binary fixture: explicitly declare the writer's existing
		// text-box type (fHaveSpt), as in the native-owned title records.
		const stream = readCompoundFileStream(generated, ['PowerPoint Document'])!.slice(), view = viewOf(stream);
		const shape = readPptSlideShapes(generated)[0]!.shapes[0]!;
		const sp = pptShapeChildren(view, readRecordOrThrow(view, shape.headerOffset)).find(record => record.recType === OA.FSP)!;
		view.setUint32(sp.dataOffset + 4, view.getUint32(sp.dataOffset + 4, true) | 0x800, true);
		const input = replaceCompoundFileStream(generated, ['PowerPoint Document'], stream);
		const document = new PptDocument(input), runs = document.slides[0]!.texts[0]!.runs;
		expect(runs.map(run => [run.start, run.end, run.text, run.directFontSizePoints])).toEqual([[0, 3, 'A😀', 20], [3, 5, 'βZ', 24]]);
		expect(() => { runs[1]!.directFontSizePoints = 30; }).toThrow(/small client anchor/);
		expect(runs[0]!.directFontSizePoints).toBe(20); expect(document.serialize()).toEqual(input);
	});
	it('preserves the previous three-argument standalone PptText constructor', () => {
		const text = new PptText({ slideId: 1, persistId: 2, textIndex: 0 }, { text: 'A', headerOffset: 8, encoding: 'utf16' }, () => {});
		expect(text.runsStatus).toBe('unsupported'); expect(text.runsDiagnostic).toMatch(/no document formatting owner/); expect(text.runs).toEqual([]);
	});
	it('rejects a reentrant document change without overwriting the inner transaction', () => {
		class ReentrantDocument extends PptDocument {
			hook?: () => void;
			protected override getBytes(): Uint8Array { const bytes = super.getBytes(); const hook = this.hook; this.hook = undefined; hook?.(); return bytes; }
		}
		const document = new ReentrantDocument(load()), first = document.slides[0]!.texts[0]!.runs[0]!, second = document.slides[1]!.texts[0]!.runs[0]!;
		document.hook = () => { second.directFontSizePoints = 30; };
		expect(() => { first.directFontSizePoints = 32; }).toThrow(/changed during character formatting edit/);
		expect(first.directFontSizePoints).toBe(28); expect(second.directFontSizePoints).toBe(30); expect(document.revision).toBe(1);
	});
	it('bounds record, aggregate style-byte and aggregate PF/CF-run work', () => {
		const input = load();
		expect(() => readPptCharacterRuns(input, { maxRecords: 1 })).toThrow(/resource limit|limit exceeded/);
		expect(readPptCharacterRuns(input, { maxStyleBytes: 1 }).every(record => record.status === 'unsupported' && /payload budget/.test(record.diagnostic!))).toBe(true);
		expect(readPptCharacterRuns(input, { maxRuns: 1 }).every(record => record.status === 'unsupported' && /run budget/.test(record.diagnostic!))).toBe(true);
		for (const maxRuns of [0, NaN, Infinity, 1.5]) expect(() => readPptCharacterRuns(input, { maxRuns })).toThrow(/Invalid.*limit/);
	});
});
