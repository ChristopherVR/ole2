import { readFileSync } from 'node:fs';
import { describe, expect, it } from 'vitest';
import { PptDocument } from '../src/ppt-document.js';
import { UnsupportedOle2EditError } from '../src/ole2-document-base.js';
import { readPptSlideShapes, pptShapeChildren } from '../src/legacy-ppt-shape-reader.js';
import { editPptShapeBounds } from '../src/legacy-ppt-shape-edit.js';
import { readCompoundFileStream, replaceCompoundFileStream } from '../src/ole2-stream-edit.js';
import { readPptSlideTexts } from '../src/legacy-ppt-text.js';
import { readRecordOrThrow } from '../src/legacy-ppt-record-stream.js';
import { buildPersistDirectory } from '../src/ppt/persist-directory.js';
import { OA } from '../src/legacy-ppt-record-types.js';
import { buildPptFile } from '../src/legacy-ppt-writer.js';

const load = () => new Uint8Array(readFileSync(new URL('./fixtures/ppt/native-text.ppt', import.meta.url)));
const viewOf = (bytes: Uint8Array) => new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
function modified(mutate: (view: DataView, shape: ReturnType<typeof readPptSlideShapes>[number]['shapes'][number]) => void): Uint8Array {
	const input = load(), stream = readCompoundFileStream(input, ['PowerPoint Document'])!, shape = readPptSlideShapes(input)[0]!.shapes[0]!;
	mutate(viewOf(stream), shape); return replaceCompoundFileStream(input, ['PowerPoint Document'], stream);
}

describe('active preserved OfficeArt shape model', () => {
	it('decodes native visible shape identities, flags, primitive kind, names and exact master-unit anchors', () => {
		const doc = new PptDocument(load());
		expect(doc.slides.map(slide => slide.shapes.map(shape => [shape.shapeId, shape.flags, shape.kind, shape.name]))).toEqual([
			[[14338, 0xa00, 'text-box', 'FixtureTitle'], [14339, 0xa00, 'text-box', 'FixtureBody'], [14340, 0xa00, 'rectangle', 'FixtureRectangle']],
			[[16386, 0xa00, 'text-box', 'FixtureUnicode'], [16387, 0xa00, 'text-box', 'FixtureWideText']],
		]);
		expect(doc.slides[0]!.shapes[0]!.bounds).toEqual({ x: 240, y: 200, width: 5200, height: 330 });
		expect(doc.slides[0]!.shapes[0]!.coordinateSpace).toBe('slide-master-units');
		expect(doc.slides[0]!.shapes[0]!.texts[0]).toBe(doc.slides[0]!.texts[0]);
		expect(doc.slides[0]!.shapes[0]!.textReference).toBe('inline');
		expect(doc.slides[0]!.shapes[2]!.mirrored).toBe(true);
		expect(Object.isFrozen(doc.slides[0]!.shapes)).toBe(true);
		expect(Object.isFrozen(doc.slides[0]!.shapes[0]!.bounds)).toBe(true);
	});
	it('edits only the fixed anchor payload and refreshes retained handles after text and geometry commits', () => {
		const input = load(), old = input.slice(), doc = new PptDocument(input), shape = doc.slides[0]!.shapes[0]!, original = shape.bounds!;
		shape.bounds = { ...original, x: 320, y: 240, width: 5280 };
		expect(shape.bounds).toEqual({ x: 320, y: 240, width: 5280, height: 330 });
		const out = doc.serialize(), stream = readCompoundFileStream(input, ['PowerPoint Document'])!, next = readCompoundFileStream(out, ['PowerPoint Document'])!;
		const anchor = readPptSlideShapes(input)[0]!.shapes[0]!.anchor!.headerOffset;
		expect(next.subarray(0, anchor + 8)).toEqual(stream.subarray(0, anchor + 8));
		expect(next.subarray(anchor + 16)).toEqual(stream.subarray(anchor + 16));
		expect(input).toEqual(old); expect(doc.revision).toBe(1);
		shape.texts[0]!.text = 'Native title updated';
		expect(shape.texts[0]!.text).toBe('Native title updated'); expect(shape.x).toBe(320);
		shape.x = 400; shape.height = 400;
		expect(doc.revision).toBe(4); expect(shape.bounds).toEqual({ x: 400, y: 240, width: 5280, height: 400 });
		const reparsed = new PptDocument(doc.serialize());
		expect(reparsed.slides[0]!.shapes[0]!.bounds).toEqual(shape.bounds);
		expect(reparsed.slides[0]!.shapes[0]!.texts[0]!.text).toBe(shape.texts[0]!.text);
	});
	it('keeps no-ops clean, including mirrored shapes, and refuses changed mirrors transactionally', () => {
		const input = load(), doc = new PptDocument(input), shape = doc.slides[0]!.shapes[2]!;
		shape.bounds = shape.bounds!; expect(doc.dirty).toBe(false);
		expect(() => { shape.x += 8; }).toThrow(/metroBlob mirror/);
		expect(doc.dirty).toBe(false); expect(doc.revision).toBe(0); expect(doc.serialize()).toEqual(input);
	});
	it.each([NaN, Infinity, 1.5, 50000])('refuses noninteger/out-of-encoding position %s without partial mutation', x => {
		const input = load(), doc = new PptDocument(input), shape = doc.slides[0]!.shapes[0]!, before = shape.bounds;
		expect(() => { shape.x = x; }).toThrow(UnsupportedOle2EditError);
		expect(shape.bounds).toEqual(before); expect(doc.serialize()).toEqual(input); expect(doc.dirty).toBe(false);
	});
	it.each([0, -1])('refuses nonpositive width %s', width => {
		const doc = new PptDocument(load()); expect(() => { doc.slides[0]!.shapes[0]!.width = width; }).toThrow(/positive/); expect(doc.dirty).toBe(false);
	});
	it('snapshots accessor geometry once and validates the actual encoded candidate', () => {
		const doc = new PptDocument(load()), shape = doc.slides[0]!.shapes[0]!, bounds = shape.bounds!;
		let reads = 0;
		shape.bounds = { x: bounds.x, y: bounds.y, get width() { return ++reads === 1 ? bounds.width + 8 : -1; }, height: bounds.height };
		expect(reads).toBe(1); expect(shape.width).toBe(bounds.width + 8);
		expect(new PptDocument(doc.serialize()).slides[0]!.shapes[0]!.width).toBe(bounds.width + 8);
	});
	it('does not overwrite a concurrent model commit made from a caller geometry accessor', () => {
		const doc = new PptDocument(load()), shape = doc.slides[0]!.shapes[0]!, before = shape.bounds!;
		expect(() => { shape.bounds = { get x() { doc.slides[0]!.texts[0]!.text = 'Native title updated'; return before.x + 8; }, y: before.y, width: before.width, height: before.height }; }).toThrow(/Document changed/);
		expect(shape.bounds).toEqual(before); expect(doc.slides[0]!.texts[0]!.text).toBe('Native title updated'); expect(doc.revision).toBe(1);
		expect(new PptDocument(doc.serialize()).slides[0]!.texts[0]!.text).toBe('Native title updated');
	});
	it.each([0xa40, 0xa80, 0xa20, 0xa10, 0xa02, 0x2a00, 0x200, 0x800])('refuses unsupported/inconsistent flags %i', flags => {
		const input = modified((view, shape) => view.setUint32(shape.headerOffset + 20, flags, true)), doc = new PptDocument(input);
		expect(() => { doc.slides[0]!.shapes[0]!.x += 8; }).toThrow(UnsupportedOle2EditError); expect(doc.serialize()).toEqual(input); expect(doc.dirty).toBe(false);
	});
	it('rejects duplicate IDs and stale expected bounds rather than selecting a record by index', () => {
		const input = load(), model = readPptSlideShapes(input), shape = model[0]!.shapes[0]!;
		expect(editPptShapeBounds(input, { slideId: model[0]!.slideId, persistId: model[0]!.persistId, shapeId: shape.shapeId, expectedBounds: { ...shape.anchor!.bounds, x: 999 }, bounds: shape.anchor!.bounds }).status).toBe('unsupported');
		const duplicate = modified((view, target) => view.setUint32(model[0]!.shapes[1]!.headerOffset + 16, target.shapeId, true)), doc = new PptDocument(duplicate);
		expect(() => { doc.slides[0]!.shapes[0]!.x += 8; }).toThrow(/ambiguous/); expect(doc.serialize()).toEqual(duplicate);
	});
	it('refuses anchor records shared through two active slide persist IDs', () => {
		const input = load(), stream = readCompoundFileStream(input, ['PowerPoint Document'])!, current = readCompoundFileStream(input, ['Current User'])!, view = viewOf(stream);
		const slides = readPptSlideTexts(input).slides, chain = buildPersistDirectory(view, viewOf(current).getUint32(16, true)), dir = readRecordOrThrow(view, chain.currentEdit.offsetPersistDirectory);
		for (let pos = dir.dataOffset; pos < dir.dataOffset + dir.recLen;) {
			const packed = view.getUint32(pos, true); pos += 4;
			for (let i = 0; i < packed >>> 20; i++, pos += 4) if ((packed & 0xfffff) + i === slides[1]!.persistId) view.setUint32(pos, chain.directory.get(slides[0]!.persistId)!, true);
		}
		// Isolate the anchor alias; duplicated live notes linkage is separately
		// invalid now that notes identities are decoded.
		const slide = readRecordOrThrow(view, chain.directory.get(slides[0]!.persistId)!);
		const atom = pptShapeChildren(view, slide).find(record => record.recType === 0x03ef)!;
		view.setUint32(atom.dataOffset + 16, 0, true);
		const shared = replaceCompoundFileStream(input, ['PowerPoint Document'], stream), doc = new PptDocument(shared), shape = doc.slides[0]!.shapes[0]!;
		expect(() => { shape.x += 8; }).toThrow(/shared/); expect(doc.serialize()).toEqual(shared); expect(doc.revision).toBe(0);
	});
	it('detects malformed FSP, anchor framing and complex property bounds', () => {
		for (const input of [
			modified((view, shape) => view.setUint32(shape.headerOffset + 4, 8, true)),
			modified((view, shape) => view.setUint16(shape.anchor!.headerOffset, 1, true)),
			modified((view, shape) => { const fopt = pptShapeChildren(view, readRecordOrThrow(view, shape.headerOffset)).find(rec => rec.recType === OA.FOPT)!; view.setUint32(fopt.dataOffset + 5 * 6 + 2, 0xffffffff, true); }),
		]) expect(() => readPptSlideShapes(input)).toThrow();
	});
	it('rejects retained handles after the physical shape identity changes', () => {
		class RevisionProbe extends PptDocument { adopt(bytes: Uint8Array) { this.commitBytes(bytes); } }
		const doc = new RevisionProbe(load()), old = doc.slides[0]!.shapes[0]!;
		const changed = modified((view, shape) => view.setUint32(shape.headerOffset + 16, shape.shapeId + 100, true));
		doc.adopt(changed); const before = doc.serialize(), revision = doc.revision;
		expect(() => old.bounds).toThrow(/identity/); expect(() => { old.x = 100; }).toThrow(/identity/);
		expect(doc.serialize()).toEqual(before); expect(doc.revision).toBe(revision);
	});
	it('inspects a large anchor without treating it as writable or guessing a missing shape-type flag', async () => {
		const input = await buildPptFile({ widthEmu: 9144000, heightEmu: 5143500, pictures: [], slides: [{ shapes: [{ kind: 'shape', spt: 1, isConnector: false, anchor: { x: 79375000, y: 1587500, w: 1270000, h: 635000 } }] }] });
		const doc = new PptDocument(input), shape = doc.slides[0]!.shapes[0]!;
		expect(shape.anchorKind).toBe('client-large'); expect(shape.bounds).toBeUndefined(); expect(shape.coordinateSpace).toBe('unknown');
		expect(shape.rawAnchorValues).toEqual([50000, 1000, 50800, 1400]);
		expect(shape.kind).toBe('unknown'); expect(shape.shapeType).toBe(1);
		expect(() => { shape.x += 8; }).toThrow(UnsupportedOle2EditError); expect(doc.serialize()).toEqual(input);
	});
	it.each(['unknown', 'duplicate', 'text-fit', 'protected'])('retains %s properties for inspection but refuses ambiguous geometry edits', variant => {
		const input = modified((view, shape) => {
			const fopt = pptShapeChildren(view, readRecordOrThrow(view, shape.headerOffset)).find(rec => rec.recType === OA.FOPT)!;
			if (variant === 'unknown') view.setUint16(fopt.dataOffset + 3 * 6, 0x147, true);
			if (variant === 'duplicate') view.setUint16(fopt.dataOffset + 3 * 6, 128, true);
			if (variant === 'text-fit') view.setUint32(fopt.dataOffset + 2 * 6 + 2, 0x60001, true);
			if (variant === 'protected') view.setUint32(fopt.dataOffset + 2, 0x40001, true);
		}), doc = new PptDocument(input);
		expect(doc.slides[0]!.shapes[0]!.geometryRefusal).toBeDefined();
		expect(() => { doc.slides[0]!.shapes[0]!.x += 8; }).toThrow(UnsupportedOle2EditError);
		expect(doc.dirty).toBe(false); expect(doc.serialize()).toEqual(input);
	});
	it.each(['maxRecords', 'maxShapes', 'maxProperties', 'maxNameChars', 'maxTextBytes'] as const)('bounds aggregate %s before materializing the model', limit => {
		expect(() => readPptSlideShapes(load(), { [limit]: 1 })).toThrow(/limit/);
		expect(() => readPptSlideShapes(load(), { [limit]: NaN })).toThrow(/Invalid shape resource limit/);
	});
	it('keeps group children in local space with real group identities and refuses geometry reconstruction', async () => {
		const input = await buildPptFile({ widthEmu: 9144000, heightEmu: 5143500, pictures: [], slides: [{ shapes: [{ kind: 'group', anchor: { x: 127000, y: 254000, w: 1270000, h: 1270000 }, childRect: { x: 0, y: 0, w: 2540000, h: 2540000 }, children: [{ kind: 'shape', spt: 1, isConnector: false, anchor: { x: 127000, y: 127000, w: 254000, h: 254000 } }] }] }] });
		const doc = new PptDocument(input), [group, child] = doc.slides[0]!.shapes;
		expect(group!.kind).toBe('group'); expect(child!.parentGroupId).toBe(group!.shapeId);
		expect(group!.parentGroupId).toBeUndefined(); expect(child!.coordinateSpace).toBe('group-child-master-units');
		expect(group!.groupBounds).toEqual({ x: 0, y: 0, width: 1600, height: 1600 });
		expect(() => { child!.x += 8; }).toThrow(/Group and child/); expect(doc.serialize()).toEqual(input);
	});
});
