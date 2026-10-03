/** Active binary PPT text and shape models backed by preserving stream edits. */
import { Ole2DocumentBase, UnsupportedOle2EditError } from './ole2-document-base.js';
import { editPptSlideText, readPptSlideTexts, type PptTextAtom } from './legacy-ppt-text.js';
import { readPptSlideShapes, type PptShapeRecord, type PptShapeBounds } from './legacy-ppt-shape-reader.js';
import { editPptShapeBounds } from './legacy-ppt-shape-edit.js';

const PPT_CAPABILITIES = Object.freeze({
	read: Object.freeze(['active-slide-ids', 'active-outline-and-inline-text', 'text-encoding', 'active-shape-ids-and-flags', 'primitive-shape-kind', 'explicit-anchor-geometry', 'validated-inline-shape-text', 'opaque-streams']),
	write: Object.freeze(['existing-text-fixed-utf16-length', 'unmirrored-top-level-small-anchor-bounds']),
	limitations: Object.freeze(['unresolved-outline-shape-text', 'rich-text-runs', 'notes', 'variable-length-text', 'OOXML-mirrored-text-and-shapes', 'group-transforms-and-geometry', 'large-anchor-writes', 'inherited-rotated-or-flipped-geometry', 'animations', 'masters', 'embedded-objects']),
});

/** An active text atom, not a shape or a fully decoded rich-text paragraph. */
export class PptText {
	readonly slideId: number;
	readonly persistId: number;
	readonly textIndex: number;
	readonly encoding: PptTextAtom['encoding'];
	/** Diagnostic stream offset; callers must not use it as an edit identity. */
	readonly headerOffset: number;
	#text: string;
	#write: (node: PptText, value: string) => void;

	constructor(identity: { slideId: number; persistId: number; textIndex: number }, atom: PptTextAtom, write: (node: PptText, value: string) => void) {
		this.slideId = identity.slideId; this.persistId = identity.persistId; this.textIndex = identity.textIndex;
		this.encoding = atom.encoding; this.headerOffset = atom.headerOffset;
		this.#text = atom.text; this.#write = write;
		Object.freeze(this);
	}
	get text(): string { return this.#text; }
	set text(value: string) {
		if (typeof value !== 'string') throw new UnsupportedOle2EditError('PPT text replacement must be a string');
		this.#write(this, value);
		// Commit the model value only after the document transaction succeeds.
		this.#text = value;
	}
	get runs(): never { throw new UnsupportedOle2EditError('PPT rich-text runs are not decoded by this adapter'); }
}

export class PptSlide {
	readonly slideId: number;
	readonly persistId: number;
	/** Active outline atoms followed by inline atoms; outline shape refs remain unresolved. */
	readonly texts: readonly PptText[];
	/** Flat OfficeArt preorder, including group descendants with parentGroupId.
	 * Canvas patriarchs/background structures remain preserved outside this list. */
	readonly shapes: readonly PptShape[];
	constructor(slideId: number, persistId: number, texts: readonly PptText[], shapes: readonly PptShape[] = []) {
		this.slideId = slideId; this.persistId = persistId;
		this.texts = Object.freeze([...texts]); this.shapes = Object.freeze([...shapes]); Object.freeze(this);
	}
	get notes(): never { throw new UnsupportedOle2EditError('PPT notes are preserved but are not decoded by this adapter'); }
}

/** A preserved OfficeArt shape. Bounds are exact master units (1/8 point),
 * local to coordinateSpace. Group children are not flattened into slide space. */
export class PptShape {
	readonly slideId: number; readonly persistId: number; readonly shapeId: number;
	readonly headerOffset: number; readonly flags: number; readonly shapeType: number;
	readonly texts: readonly PptText[];
	#read: () => PptShapeRecord;
	#write: (bounds: PptShapeBounds) => void;
	constructor(identity: { slideId: number; persistId: number }, record: PptShapeRecord, texts: readonly PptText[], read: () => PptShapeRecord, write: (bounds: PptShapeBounds) => void) {
		this.slideId = identity.slideId; this.persistId = identity.persistId; this.shapeId = record.shapeId;
		this.headerOffset = record.headerOffset; this.flags = record.flags; this.shapeType = record.shapeType;
		this.texts = Object.freeze([...texts]); this.#read = read; this.#write = write; Object.freeze(this);
	}
	get kind(): PptShapeRecord['kind'] { return this.#read().kind; }
	get name(): string | undefined { return this.#read().name; }
	get coordinateSpace(): PptShapeRecord['coordinateSpace'] { return this.#read().coordinateSpace; }
	get parentGroupId(): number | undefined { return this.#read().parentGroupId; }
	get anchorKind(): NonNullable<PptShapeRecord['anchor']>['kind'] | undefined { return this.#read().anchor?.kind; }
	get mirrored(): boolean { return this.#read().mirrored; }
	get rotation(): number { return this.#read().rotation; }
	get textReference(): PptShapeRecord['textReference'] { return this.#read().textReference; }
	get geometryRefusal(): string | undefined { return this.#read().geometryRefusal; }
	get groupBounds(): Readonly<PptShapeBounds> | undefined { const bounds = this.#read().groupBounds; return bounds && Object.freeze({ ...bounds }); }
	get bounds(): Readonly<PptShapeBounds> | undefined { const bounds = this.#read().anchor?.bounds; return bounds && Object.freeze({ ...bounds }); }
	set bounds(value: PptShapeBounds) { this.#write(value); }
	#requiredBounds(): PptShapeBounds { const bounds = this.bounds; if (!bounds) throw new UnsupportedOle2EditError('Shape has no decoded explicit anchor'); return bounds; }
	get x(): number { return this.#requiredBounds().x; }
	set x(value: number) { this.bounds = { ...this.#requiredBounds(), x: value }; }
	get y(): number { return this.#requiredBounds().y; }
	set y(value: number) { this.bounds = { ...this.#requiredBounds(), y: value }; }
	get width(): number { return this.#requiredBounds().width; }
	set width(value: number) { this.bounds = { ...this.#requiredBounds(), width: value }; }
	get height(): number { return this.#requiredBounds().height; }
	set height(value: number) { this.bounds = { ...this.#requiredBounds(), height: value }; }
}

/** Indexed active slides with transactional fixed-length text and guarded small
 * anchor edits. Opaque properties, formatting, notes and unknown records stay
 * in the CFB. Shape text links are exposed only for validated inline atoms;
 * unresolved outline references are not inferred from positional indexes. */
export class PptDocument extends Ole2DocumentBase {
	get kind(): 'ppt' { return 'ppt'; }
	get capabilities(): typeof PPT_CAPABILITIES { return PPT_CAPABILITIES; }
	readonly #slides: readonly PptSlide[];
	readonly #unsupported: readonly string[];
	#shapeCache: { revision: number; slides: ReturnType<typeof readPptSlideShapes> };
	get slides(): readonly PptSlide[] { return this.#slides; }
	get unsupported(): readonly string[] { return this.#unsupported; }

	constructor(input: Uint8Array) {
		super(input);
		const shapeModel = readPptSlideShapes(this.getBytes());
		const model = readPptSlideTexts(this.getBytes());
		this.#shapeCache = { revision: this.revision, slides: shapeModel };
		this.#unsupported = Object.freeze(['formatting', 'group-transforms', 'pictures-content', 'masters', 'notes', 'animations', 'embedded-objects', 'unresolved-outline-shape-text', 'variable-length-text']);
		this.#slides = Object.freeze(model.slides.map((slide, index) => {
			const texts = slide.texts.map((atom, textIndex) => new PptText({ slideId: slide.slideId, persistId: slide.persistId, textIndex }, atom, (node, value) => this.#edit(node, value)));
			const shapes = shapeModel[index]!.shapes.map(record => new PptShape(slide, record, record.textIndexes.map(textIndex => texts[textIndex]!),
				() => this.#readShape(slide.slideId, slide.persistId, record), value => this.#editShape(slide.slideId, slide.persistId, record, value)));
			return new PptSlide(slide.slideId, slide.persistId, texts, shapes);
		}));
	}

	#readShape(slideId: number, persistId: number, original: PptShapeRecord): PptShapeRecord {
		if (this.#shapeCache.revision !== this.revision) this.#shapeCache = { revision: this.revision, slides: readPptSlideShapes(this.getBytes()) };
		const slides = this.#shapeCache.slides.filter(slide => slide.slideId === slideId && slide.persistId === persistId);
		const matches = slides.length === 1 ? slides[0]!.shapes.filter(shape => shape.shapeId === original.shapeId) : [];
		const current = matches[0];
		if (matches.length !== 1 || !current || current.headerOffset !== original.headerOffset || current.flags !== original.flags || current.shapeType !== original.shapeType)
			throw new UnsupportedOle2EditError('PPT shape identity no longer matches or is ambiguous');
		return current;
	}

	#editShape(slideId: number, persistId: number, original: PptShapeRecord, value: PptShapeBounds): void {
		const revision = this.revision;
		const current = this.#readShape(slideId, persistId, original);
		if (!current.anchor) throw new UnsupportedOle2EditError('PPT shape has no decoded anchor');
		const result = editPptShapeBounds(this.getBytes(), { slideId, persistId, shapeId: current.shapeId, expectedHeaderOffset: current.headerOffset, expectedBounds: current.anchor.bounds, bounds: value });
		if (this.revision !== revision) throw new UnsupportedOle2EditError('Document changed during the geometry edit');
		if (result.status === 'unsupported') throw new UnsupportedOle2EditError(result.reason);
		if (result.status === 'edited') this.commitBytes(result.bytes);
	}

	#edit(node: PptText, value: string): void {
		const bytes = this.getBytes(), current = readPptSlideTexts(bytes);
		const slideIndex = current.slides.findIndex(slide => slide.slideId === node.slideId && slide.persistId === node.persistId);
		const atom = current.slides[slideIndex]?.texts[node.textIndex];
		if (!atom || atom.headerOffset !== node.headerOffset || atom.encoding !== node.encoding || atom.text !== node.text)
			throw new UnsupportedOle2EditError('PPT text identity or expected text no longer matches');
		if (value !== atom.text && current.slides.reduce((count, slide) => count + slide.texts.filter(text => text.headerOffset === atom.headerOffset).length, 0) !== 1)
			throw new UnsupportedOle2EditError('PPT text atom is shared by multiple active locations');
		const result = editPptSlideText(bytes, { slideIndex, textIndex: node.textIndex, expectedText: node.text, text: value });
		if (result.status === 'unsupported') throw new UnsupportedOle2EditError(result.reason);
		if (result.status === 'edited') this.commitBytes(result.bytes);
	}
}
