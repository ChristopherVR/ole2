/** Fixed-record geometry edits; no drawing reconstruction or property loss. */
import { readCompoundFileStream, replaceCompoundFileStream } from './ole2-stream-edit.js';
import { readPptSlideShapes, type PptShapeBounds } from './legacy-ppt-shape-reader.js';
import { readRecordOrThrow, PptParseError } from './legacy-ppt-record-stream.js';
import { buildPersistDirectory, parseUserEditAtom } from './ppt/persist-directory.js';

export type PptShapeEditResult =
	| { status: 'edited' | 'unchanged'; bytes: Uint8Array }
	| { status: 'unsupported'; bytes: Uint8Array; reason: string };

export function equalPptShapeBounds(a: PptShapeBounds, b: PptShapeBounds): boolean {
	return a.x === b.x && a.y === b.y && a.width === b.width && a.height === b.height;
}

/** The entire slide must not also belong to an opaque live persist object or
 * save-chain metadata. Inspecting other visible shapes alone misses aliases. */
function slideOwnershipRefusal(input: Uint8Array, stream: Uint8Array, persistId: number): string | undefined {
	const user = readCompoundFileStream(input, ['Current User']);
	if (!user) throw new PptParseError('Missing Current User stream');
	const view = new DataView(stream.buffer, stream.byteOffset, stream.byteLength), uv = new DataView(user.buffer, user.byteOffset, user.byteLength);
	const cu = readRecordOrThrow(uv, 0), editOffset = uv.getUint32(cu.dataOffset + 8, true);
	const { directory } = buildPersistDirectory(view, editOffset), ownedOffset = directory.get(persistId);
	if (ownedOffset === undefined) return 'Missing active slide persist object';
	const owned = readRecordOrThrow(view, ownedOffset), start = owned.headerOffset, end = owned.dataOffset + owned.recLen;
	if (end > view.byteLength) return 'Unbounded slide persist object';
	const overlaps = (offset: number): boolean => {
		const other = readRecordOrThrow(view, offset), otherEnd = other.dataOffset + other.recLen;
		if (otherEnd > view.byteLength) throw new PptParseError('Unbounded live persist object prevents safe geometry edit');
		return start < otherEnd && other.headerOffset < end;
	};
	for (const [id, offset] of directory) {
		if (id !== persistId && overlaps(offset)) return 'Slide container overlaps another live persist object';
	}
	const visited = new Set<number>();
	for (let offset = editOffset; offset;) {
		if (visited.has(offset) || visited.size >= 10000) return 'Invalid geometry save-history ownership';
		visited.add(offset);
		const edit = parseUserEditAtom(view, offset);
		if (overlaps(offset) || overlaps(edit.offsetPersistDirectory)) return 'Slide container overlaps save-history metadata';
		offset = edit.offsetLastEdit;
	}
	return undefined;
}

/** Bounds use exact master units (1/8 point). Only existing, untransformed,
 * unmirrored rectangle/text-box small client anchors are currently writable. */
export function editPptShapeBounds(input: Uint8Array, edit: {
	slideId: number; persistId: number; shapeId: number; expectedBounds: PptShapeBounds;
	bounds: PptShapeBounds; expectedHeaderOffset?: number;
}): PptShapeEditResult {
	const reject = (reason: string): PptShapeEditResult => ({ status: 'unsupported', bytes: input, reason });
	try {
		const { slideId, persistId, shapeId, expectedHeaderOffset, bounds: supplied, expectedBounds: suppliedExpected } = edit;
		// Caller objects can contain accessors. Snapshot each field once so
		// validation and encoding always act on the same values.
		if (!supplied || !suppliedExpected) return reject('Missing replacement or expected geometry');
		const next = { x: supplied.x, y: supplied.y, width: supplied.width, height: supplied.height };
		const expected = { x: suppliedExpected.x, y: suppliedExpected.y, width: suppliedExpected.width, height: suppliedExpected.height };
		const model = readPptSlideShapes(input), slides = model.filter(slide => slide.slideId === slideId && slide.persistId === persistId);
		const matches = slides.length === 1 ? slides[0]!.shapes.filter(shape => shape.shapeId === shapeId) : [];
		if (matches.length !== 1) return reject('Shape identity is missing or ambiguous');
		const shape = matches[0]!;
		if (!shape.anchor?.bounds || !equalPptShapeBounds(shape.anchor.bounds, expected) || (expectedHeaderOffset !== undefined && shape.headerOffset !== expectedHeaderOffset))
			return reject('Shape identity or expected geometry no longer matches');
		if (![next.x, next.y, next.width, next.height].every(Number.isSafeInteger) || next.width <= 0 || next.height <= 0)
			return reject('Geometry must use integer master units and positive extents');
		if (equalPptShapeBounds(next, shape.anchor.bounds)) return { status: 'unchanged', bytes: input };
		if (shape.geometryRefusal) return reject(shape.geometryRefusal);
		const anchor = shape.anchor.headerOffset, end = anchor + 16;
		// Reject shared/overlapping physical records across active slide objects.
		for (const slide of model) for (const other of slide.shapes) {
			if (other === shape) continue;
			if ((other.headerOffset <= anchor && other.recordEnd > anchor) || (other.anchor && other.anchor.headerOffset < end && other.anchor.headerOffset + (other.anchor.kind === 'client-small' ? 16 : 24) > anchor))
				return reject('Anchor is shared by multiple active shape locations');
		}
		const edges = [next.y, next.x, next.x + next.width, next.y + next.height];
		if (!edges.every(value => Number.isSafeInteger(value) && value >= -32768 && value <= 32767)) return reject('Geometry exceeds the existing small anchor encoding');
		const stream = readCompoundFileStream(input, ['PowerPoint Document'])!;
		const ownership = slideOwnershipRefusal(input, stream, persistId);
		if (ownership) return reject(ownership);
		const out = stream.slice(), view = new DataView(out.buffer);
		for (let i = 0; i < edges.length; i++) view.setInt16(anchor + 8 + i * 2, edges[i]!, true);
		const bytes = replaceCompoundFileStream(input, ['PowerPoint Document'], out);
		if (bytes === input) return reject('Compound stream cannot be safely replaced');
		const candidate = readPptSlideShapes(bytes).filter(slide => slide.slideId === slideId && slide.persistId === persistId).flatMap(slide => slide.shapes.filter(item => item.shapeId === shapeId));
		if (candidate.length !== 1 || candidate[0]!.headerOffset !== shape.headerOffset || !candidate[0]!.anchor?.bounds || !equalPptShapeBounds(candidate[0]!.anchor!.bounds!, next))
			return reject('Edited shape failed candidate validation');
		return { status: 'edited', bytes };
	} catch (error) { return reject(error instanceof Error ? error.message : 'Malformed PPT shape'); }
}
