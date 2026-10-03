/** Bounded active OfficeArt shape inspection. Opaque records stay in source bytes. */
import { readCompoundFileStream } from './ole2-stream-edit.js';
import { readPptSlideTexts, PptTextError } from './legacy-ppt-text.js';
import { readRecordOrThrow, type PptRecord } from './legacy-ppt-record-stream.js';
import { buildPersistDirectory } from './ppt/persist-directory.js';
import { OA, RT, HEADER_TOKEN_ENCRYPTED } from './legacy-ppt-record-types.js';

export interface PptShapeReadLimits { maxRecords?: number; maxShapes?: number; maxProperties?: number; maxNameChars?: number; maxTextBytes?: number }
const DEFAULT_LIMITS = { maxRecords: 100000, maxShapes: 10000, maxProperties: 100000, maxNameChars: 1000000, maxTextBytes: 16 * 1024 * 1024 };
// Only properties whose anchor interaction is understood for this increment.
// Unknown properties remain readable and byte-preserved, but block geometry edits.
const GEOMETRY_PROPERTY_IDS = new Set([4, 127, 128, 191, 385, 447, 448, 459, 511, 896, 959]);

/** Exact master units: 576 per inch, or 8 per PowerPoint point. */
export interface PptShapeBounds { x: number; y: number; width: number; height: number }
export interface PptShapeRecord {
	shapeId: number; shapeType: number; flags: number;
	kind: 'rectangle' | 'ellipse' | 'line' | 'picture' | 'text-box' | 'group' | 'unknown';
	headerOffset: number; recordEnd: number; parentGroupId?: number;
	anchor?: { kind: 'client-small' | 'client-large' | 'child'; headerOffset: number; bounds?: PptShapeBounds; rawValues?: number[] };
	coordinateSpace: 'slide-master-units' | 'group-child-master-units' | 'unknown';
	groupBounds?: PptShapeBounds; name?: string; rotation: number; mirrored: boolean;
	textIndexes: number[]; textReference: 'inline' | 'outline' | 'none' | 'unresolved';
	propertyIds: number[]; geometryRefusal?: string;
}
export interface PptSlideShapes { slideId: number; persistId: number; shapes: PptShapeRecord[] }

export function pptShapeChildren(view: DataView, parent: PptRecord): PptRecord[] {
	const records: PptRecord[] = [], end = parent.dataOffset + parent.recLen;
	if (end > view.byteLength) throw new PptTextError('corrupt', 'Shape container exceeds stream');
	for (let pos = parent.dataOffset; pos < end;) {
		if (records.length >= DEFAULT_LIMITS.maxRecords) throw new PptTextError('corrupt', 'PPT shape resource limit exceeded: child records');
		if (end - pos < 8) throw new PptTextError('corrupt', 'Truncated shape child header');
		const rec = readRecordOrThrow(view, pos);
		if (rec.dataOffset + rec.recLen > end) throw new PptTextError('corrupt', 'Shape child exceeds container');
		records.push(rec); pos = rec.dataOffset + rec.recLen;
	}
	return records;
}

function bounds(view: DataView, rec: PptRecord): PptShapeBounds {
	const d = rec.dataOffset;
	// Only small ClientAnchor and OfficeArt child/group RECT layouts call this.
	// Large ClientAnchor axis order is kept undecoded pending retained evidence.
	const [x, y, right, bottom] = rec.recLen === 8
		? [view.getInt16(d + 2, true), view.getInt16(d, true), view.getInt16(d + 4, true), view.getInt16(d + 6, true)]
		: [view.getInt32(d, true), view.getInt32(d + 4, true), view.getInt32(d + 8, true), view.getInt32(d + 12, true)];
	return { x: x!, y: y!, width: right! - x!, height: bottom! - y! };
}

/** Decode active slide shapes, excluding canvas patriarchs and backgrounds.
 * Group child coordinates remain local; no inferred global transform is made.
 * Existing text atoms are linked by physical identity or validated outline refs. */
export function readPptSlideShapes(input: Uint8Array, limits: PptShapeReadLimits = {}): PptSlideShapes[] {
	const budgets = { ...DEFAULT_LIMITS, ...limits };
	if (!Object.values(budgets).every(value => Number.isSafeInteger(value) && value > 0)) throw new PptTextError('corrupt', 'Invalid shape resource limit');
	const stream = readCompoundFileStream(input, ['PowerPoint Document']), current = readCompoundFileStream(input, ['Current User']);
	if (!stream || !current || current.length < 20) throw new PptTextError('corrupt', 'Missing or malformed PowerPoint streams');
	const view = new DataView(stream.buffer, stream.byteOffset, stream.byteLength), cv = new DataView(current.buffer, current.byteOffset, current.byteLength);
	if (cv.getUint32(12, true) === HEADER_TOKEN_ENCRYPTED) throw new PptTextError('encrypted', 'Encrypted PPT shapes are unsupported');
	const chain = buildPersistDirectory(view, cv.getUint32(16, true), { maxEntries: budgets.maxRecords });
	const documentOffset = chain.directory.get(chain.currentEdit.docPersistIdRef);
	if (documentOffset === undefined) throw new PptTextError('corrupt', 'Missing active document');
	const documentRecord = readRecordOrThrow(view, documentOffset);
	const slideOffsets: number[] = [], activeLists = pptShapeChildren(view, documentRecord).filter(record => record.recType === RT.SlideListWithText && record.recInstance === 0);
	for (const list of activeLists) for (const record of pptShapeChildren(view, list)) if (record.recType === RT.SlidePersistAtom) {
		if (record.recLen !== 20) throw new PptTextError('corrupt', 'Invalid slide persist atom');
		const offset = chain.directory.get(view.getUint32(record.dataOffset, true));
		if (offset === undefined) throw new PptTextError('corrupt', 'Missing active slide');
		slideOffsets.push(offset);
		if (slideOffsets.length > budgets.maxShapes) throw new PptTextError('corrupt', 'PPT shape resource limit exceeded: slide references');
	}
	// Count every active reference, including aliases, before decoding strings
	// or materializing shape arrays. Repeated offsets must consume the budget.
	let recordCount = 0, shapeCount = 0, propertyCount = 0, nameChars = 0, textBytes = 0;
	const preflight = [...activeLists, ...slideOffsets.map(offset => readRecordOrThrow(view, offset))].map(rec => ({ rec, depth: 0 }));
	while (preflight.length) {
		const { rec, depth } = preflight.pop()!;
		if (++recordCount > budgets.maxRecords || depth > 128) throw new PptTextError('corrupt', 'PPT shape resource limit exceeded: records or depth');
		if (rec.recType === OA.SpContainer && ++shapeCount > budgets.maxShapes) throw new PptTextError('corrupt', 'PPT shape resource limit exceeded: shapes');
		if (rec.recType === RT.TextCharsAtom || rec.recType === RT.TextBytesAtom) {
			textBytes += rec.recLen;
			if (textBytes > budgets.maxTextBytes) throw new PptTextError('corrupt', 'PPT shape resource limit exceeded: text bytes');
		}
		if ([OA.FOPT, OA.TertiaryFOPT, 0xf121].includes(rec.recType)) {
			propertyCount += rec.recInstance;
			if (propertyCount > budgets.maxProperties) throw new PptTextError('corrupt', 'PPT shape resource limit exceeded: properties');
			if (rec.recInstance * 6 > rec.recLen || rec.dataOffset + rec.recLen > view.byteLength) throw new PptTextError('corrupt', 'Invalid shape property table');
			for (let i = 0; i < rec.recInstance; i++) if ((view.getUint16(rec.dataOffset + i * 6, true) & 0x3fff) === 896) {
				nameChars += Math.ceil(view.getUint32(rec.dataOffset + i * 6 + 2, true) / 2);
				if (nameChars > budgets.maxNameChars) throw new PptTextError('corrupt', 'PPT shape resource limit exceeded: name characters');
			}
		}
		if (rec.recVer === 15 || rec.recType === OA.ClientTextbox) for (const child of pptShapeChildren(view, rec)) preflight.push({ rec: child, depth: depth + 1 });
	}
	const textModel = readPptSlideTexts(input);
	return textModel.slides.map(slide => {
		const slideRecord = readRecordOrThrow(view, chain.directory.get(slide.persistId)!);
		const shapes: PptShapeRecord[] = [], stack: Array<{ rec: PptRecord; depth: number; group?: number }> = [{ rec: slideRecord, depth: 0 }];
		while (stack.length) {
			const item = stack.pop()!;
			if (item.depth > 128) throw new PptTextError('corrupt', 'Shape nesting exceeds limit');
			const children = pptShapeChildren(view, item.rec);
			if (item.rec.recType === OA.SpContainer) {
				const fsps = children.filter(rec => rec.recType === OA.FSP);
				if (fsps.length !== 1 || fsps[0]!.recVer !== 2 || fsps[0]!.recLen !== 8)
					throw new PptTextError('corrupt', 'Missing or ambiguous OfficeArtFSP');
				const fsp = fsps[0]!, shapeId = view.getUint32(fsp.dataOffset, true), flags = view.getUint32(fsp.dataOffset + 4, true);
				if (!shapeId) throw new PptTextError('corrupt', 'Zero shape identifier');
				if (flags & (0x4 | 0x400)) continue;
				const shape: PptShapeRecord = {
					shapeId, shapeType: fsp.recInstance, flags, headerOffset: item.rec.headerOffset, recordEnd: item.rec.dataOffset + item.rec.recLen,
					kind: flags & 1 ? 'group' : flags & 0x800 ? ({ 1: 'rectangle', 3: 'ellipse', 20: 'line', 75: 'picture', 202: 'text-box' } as const)[fsp.recInstance as 1] ?? 'unknown' : 'unknown',
					parentGroupId: item.group, coordinateSpace: 'unknown', rotation: 0, mirrored: false, propertyIds: [], textIndexes: [], textReference: 'none',
				};
				let propertyAmbiguity = false, unknownProperties = false, protectedGeometry = false, unknownTextFit = false;
				const seenProperties = new Set<number>();
				const anchors = children.filter(rec => rec.recType === OA.ClientAnchor || rec.recType === OA.ChildAnchor);
				if (anchors.length === 1) {
					const rec = anchors[0]!;
					if (rec.recVer !== 0 || rec.recInstance !== 0 || (rec.recType === OA.ClientAnchor ? rec.recLen !== 8 && rec.recLen !== 16 : rec.recLen !== 16))
						throw new PptTextError('corrupt', 'Invalid OfficeArt anchor');
					const largeClient = rec.recType === OA.ClientAnchor && rec.recLen === 16;
					shape.anchor = { kind: rec.recType === OA.ChildAnchor ? 'child' : rec.recLen === 8 ? 'client-small' : 'client-large', headerOffset: rec.headerOffset,
						bounds: largeClient ? undefined : bounds(view, rec), rawValues: largeClient ? Array.from({ length: 4 }, (_, i) => view.getInt32(rec.dataOffset + i * 4, true)) : undefined };
					shape.coordinateSpace = largeClient ? 'unknown' : rec.recType === OA.ChildAnchor ? 'group-child-master-units' : item.group !== undefined || flags & 2 ? 'unknown' : 'slide-master-units';
				}
				const group = children.find(rec => rec.recType === OA.FSPGR);
				if (group) {
					if (group.recVer !== 1 || group.recLen !== 16) throw new PptTextError('corrupt', 'Invalid group coordinate space');
					shape.groupBounds = bounds(view, group);
				}
				for (const rec of children.filter(rec => rec.recType === OA.FOPT || rec.recType === OA.TertiaryFOPT || rec.recType === 0xf121)) {
					if (rec.recVer !== 3 || rec.recInstance * 6 > rec.recLen) throw new PptTextError('corrupt', 'Invalid shape property table');
					let complexOffset = rec.dataOffset + rec.recInstance * 6;
					for (let i = 0; i < rec.recInstance; i++) {
						const at = rec.dataOffset + i * 6, packed = view.getUint16(at, true), id = packed & 0x3fff, value = view.getUint32(at + 2, true);
						shape.propertyIds.push(id);
						if (seenProperties.has(id)) propertyAmbiguity = true;
						seenProperties.add(id);
						if (!GEOMETRY_PROPERTY_IDS.has(id) || ((packed & 0x4000) && !(id === 896 && (packed & 0x8000))) || ((packed & 0x8000) && id !== 896)) unknownProperties = true;
						if (id === 127 && (value & 0xffff)) protectedGeometry = true;
						if (id === 191 && value !== 0x60002) unknownTextFit = true;
						if (id === 0x3a9) shape.mirrored = true;
						if (id === 4) shape.rotation = view.getInt32(at + 2, true) / 65536;
						if (packed & 0x8000) {
							if (value > rec.dataOffset + rec.recLen - complexOffset) throw new PptTextError('corrupt', 'Shape complex property exceeds record');
							if (id === 0x380 && value >= 2 && value % 2 === 0 && view.getUint16(complexOffset + value - 2, true) === 0) {
								let name = ''; for (let p = 0; p < value - 2; p += 2) name += String.fromCharCode(view.getUint16(complexOffset + p, true)); shape.name = name;
							}
							complexOffset += value;
						}
					}
					if (complexOffset !== rec.dataOffset + rec.recLen) propertyAmbiguity = true;
				}
				shape.textIndexes = slide.texts.flatMap((atom, index) => atom.headerOffset >= item.rec.dataOffset && atom.headerOffset < shape.recordEnd ? [index] : []);
				if (shape.textIndexes.length) shape.textReference = 'inline';
				// Outline references are retained, but not guessed from atom index.
				for (const textbox of children.filter(rec => rec.recType === OA.ClientTextbox))
					if (pptShapeChildren(view, textbox).some(rec => rec.recType === RT.OutlineTextRefAtom)) shape.textReference = 'unresolved';
				if (shape.mirrored) shape.geometryRefusal = 'Shape has an OOXML metroBlob mirror';
				else if (propertyAmbiguity) shape.geometryRefusal = 'Shape property duplicates or unparsed payload are ambiguous';
				else if (unknownProperties || protectedGeometry || unknownTextFit) shape.geometryRefusal = 'Shape has unvalidated or protected geometry-related properties';
				else if (item.group !== undefined || flags & (1 | 2)) shape.geometryRefusal = 'Group and child shape transforms are unsupported';
				else if (flags & ~0xa00 || !(flags & 0x800) || shape.rotation || shape.propertyIds.includes(0x301)) shape.geometryRefusal = 'Shape has unsupported flags or inherited/rotated geometry';
				else if (!(flags & 0x200) || anchors.length !== 1 || shape.anchor?.kind !== 'client-small') shape.geometryRefusal = 'Only one explicit small client anchor is editable';
				else if (shape.kind !== 'rectangle' && shape.kind !== 'text-box') shape.geometryRefusal = 'Primitive shape kind is not native-validated for geometry edits';
				else if (!shape.anchor.bounds || shape.anchor.bounds.width <= 0 || shape.anchor.bounds.height <= 0) shape.geometryRefusal = 'Nonpositive or undecoded shape extents are unsupported';
				shapes.push(shape); continue;
			}
			let parentGroup = item.group;
			if (item.rec.recType === OA.SpgrContainer) {
				const head = children[0];
				if (!head || head.recType !== OA.SpContainer) throw new PptTextError('corrupt', 'Group has no shape container');
				const fsp = pptShapeChildren(view, head).find(rec => rec.recType === OA.FSP);
				if (!fsp || fsp.recLen !== 8) throw new PptTextError('corrupt', 'Group has no FSP');
				if (!(view.getUint32(fsp.dataOffset + 4, true) & 4)) parentGroup = view.getUint32(fsp.dataOffset, true);
			}
			for (let i = children.length - 1; i >= 0; i--) {
				const rec = children[i]!;
				if (rec.recVer === 15 || rec.recType === OA.ClientTextbox) stack.push({ rec, depth: item.depth + 1, group: item.rec.recType === OA.SpgrContainer && i === 0 ? item.group : parentGroup });
			}
		}
		return { slideId: slide.slideId, persistId: slide.persistId, shapes };
	});
}
