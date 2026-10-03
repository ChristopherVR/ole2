/** Fixed-record geometry edits; no drawing reconstruction or property loss. */
import { readCompoundFileStream, replaceCompoundFileStream } from './ole2-stream-edit.js';
import { readPptSlideShapes, type PptShapeBounds } from './legacy-ppt-shape-reader.js';

export type PptShapeEditResult =
	| { status: 'edited' | 'unchanged'; bytes: Uint8Array }
	| { status: 'unsupported'; bytes: Uint8Array; reason: string };

export function equalPptShapeBounds(a: PptShapeBounds, b: PptShapeBounds): boolean {
	return a.x === b.x && a.y === b.y && a.width === b.width && a.height === b.height;
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
		if (!shape.anchor || !equalPptShapeBounds(shape.anchor.bounds, expected) || (expectedHeaderOffset !== undefined && shape.headerOffset !== expectedHeaderOffset))
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
		const stream = readCompoundFileStream(input, ['PowerPoint Document'])!, out = stream.slice(), view = new DataView(out.buffer);
		for (let i = 0; i < edges.length; i++) view.setInt16(anchor + 8 + i * 2, edges[i]!, true);
		const bytes = replaceCompoundFileStream(input, ['PowerPoint Document'], out);
		if (bytes === input) return reject('Compound stream cannot be safely replaced');
		const candidate = readPptSlideShapes(bytes).filter(slide => slide.slideId === slideId && slide.persistId === persistId).flatMap(slide => slide.shapes.filter(item => item.shapeId === shapeId));
		if (candidate.length !== 1 || candidate[0]!.headerOffset !== shape.headerOffset || !candidate[0]!.anchor || !equalPptShapeBounds(candidate[0]!.anchor!.bounds, next))
			return reject('Edited shape failed candidate validation');
		return { status: 'edited', bytes };
	} catch (error) { return reject(error instanceof Error ? error.message : 'Malformed PPT shape'); }
}
