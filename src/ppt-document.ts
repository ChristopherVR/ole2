/** Active binary PPT text and shape models backed by preserving stream edits. */
import { Ole2DocumentBase, UnsupportedOle2EditError } from './ole2-document-base.js';
import { editPptSlideText, readPptSlideTexts, type PptTextAtom } from './legacy-ppt-text.js';
import { readPptSlideShapes, type PptShapeRecord, type PptShapeBounds } from './legacy-ppt-shape-reader.js';
import { editPptShapeBounds } from './legacy-ppt-shape-edit.js';
import { readPptSlideNotes, editPptNotesText, type PptSlideNotes, type PptNoteTextAtom } from './legacy-ppt-notes.js';
import { readPptCharacterRuns, editPptCharacterFontSize, type PptTextRunRecord, type PptCharacterRunRecord } from './legacy-ppt-text-style.js';

const PPT_CAPABILITIES = Object.freeze({
	read: Object.freeze(['active-slide-ids', 'active-outline-and-inline-text', 'text-encoding', 'direct-character-run-spans-and-font-size', 'active-shape-ids-and-flags', 'primitive-shape-kind', 'explicit-anchor-geometry', 'validated-inline-shape-text', 'active-notes-identities-and-inline-text', 'opaque-streams']),
	write: Object.freeze(['existing-text-fixed-utf16-length', 'existing-literal-inline-text-box-font-size', 'unmirrored-top-level-small-anchor-bounds', 'existing-notes-body-fixed-utf16-length']),
	limitations: Object.freeze(['unresolved-outline-shape-text', 'inherited-character-formatting', 'run-splitting-and-style-insertion', 'character-properties-except-existing-font-size', 'outline-and-unsupported-shape-font-size-edits', 'notes-creation-and-inherited-fields', 'variable-length-text', 'OOXML-mirrored-text-and-shapes', 'group-transforms-and-geometry', 'large-anchor-writes', 'inherited-rotated-or-flipped-geometry', 'animations', 'masters', 'embedded-objects']),
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
	#readRuns: (node: PptText) => PptTextRunRecord;
	#writeRun: (node: PptText, run: PptCharacterRunRecord, value: number | undefined) => void;
	#runs: readonly PptCharacterRun[] | undefined;

	constructor(identity: { slideId: number; persistId: number; textIndex: number }, atom: PptTextAtom, write: (node: PptText, value: string) => void,
		readRuns?: (node: PptText) => PptTextRunRecord, writeRun?: (node: PptText, run: PptCharacterRunRecord, value: number | undefined) => void) {
		this.slideId = identity.slideId; this.persistId = identity.persistId; this.textIndex = identity.textIndex;
		this.encoding = atom.encoding; this.headerOffset = atom.headerOffset;
		this.#text = atom.text; this.#write = write;
		this.#readRuns = readRuns ?? (() => ({ ...identity, headerOffset: atom.headerOffset, text: this.#text, status: 'unsupported', diagnostic: 'Text has no document formatting owner', runs: [] }));
		this.#writeRun = writeRun ?? (() => { throw new UnsupportedOle2EditError('Text has no document formatting owner'); });
		Object.freeze(this);
	}
	get text(): string { return this.#text; }
	set text(value: string) {
		if (typeof value !== 'string') throw new UnsupportedOle2EditError('PPT text replacement must be a string');
		this.#write(this, value);
		// Commit the model value only after the document transaction succeeds.
		this.#text = value;
	}
	get runsStatus(): PptTextRunRecord['status'] { return this.#readRuns(this).status; }
	get runsDiagnostic(): string | undefined { return this.#readRuns(this).diagnostic; }
	/** Direct character exceptions only; inherited/effective formatting is unresolved. */
	get runs(): readonly PptCharacterRun[] {
		if (!this.#runs) {
			const model = this.#readRuns(this);
			this.#runs = Object.freeze(model.runs.map(run => new PptCharacterRun(run, () => {
				const current = this.#readRuns(this);
				const matches = current.status === 'decoded' ? current.runs.filter(item => item.runOffset === run.runOffset && item.styleHeaderOffset === run.styleHeaderOffset && item.start === run.start && item.end === run.end && item.rawCount === run.rawCount && item.mask === run.mask && item.fontSizeOffset === run.fontSizeOffset) : [];
				if (matches.length !== 1) throw new UnsupportedOle2EditError('PPT character run identity no longer matches or is ambiguous');
				return { text: current.text, run: matches[0]! };
			}, value => this.#writeRun(this, run, value))));
		}
		return this.#runs;
	}
}

/** Stable UTF-16 span of an existing direct character-format run. The native
 * implicit paragraph mark is preserved; it is excluded from start/end/text. */
export class PptCharacterRun {
	readonly start: number; readonly end: number;
	readonly styleHeaderOffset: number; readonly runOffset: number;
	#read: () => { text: string; run: PptCharacterRunRecord };
	#write: (value: number | undefined) => void;
	constructor(record: PptCharacterRunRecord, read: () => { text: string; run: PptCharacterRunRecord }, write: (value: number | undefined) => void) {
		this.start = record.start; this.end = record.end; this.styleHeaderOffset = record.styleHeaderOffset; this.runOffset = record.runOffset;
		this.#read = read; this.#write = write; Object.freeze(this);
	}
	get text(): string { return this.#read().text.slice(this.start, this.end); }
	get directFontSizePoints(): number | undefined { return this.#read().run.directFontSizePoints; }
	set directFontSizePoints(value: number | undefined) {
		if (value !== undefined && typeof value !== 'number') throw new UnsupportedOle2EditError('PPT direct font size must be numeric');
		this.#write(value);
	}
}

export class PptSlide {
	readonly slideId: number;
	readonly persistId: number;
	/** Active outline atoms followed by inline atoms; outline shape refs remain unresolved. */
	readonly texts: readonly PptText[];
	/** Flat OfficeArt preorder, including group descendants with parentGroupId.
	 * Canvas patriarchs/background structures remain preserved outside this list. */
	readonly shapes: readonly PptShape[];
	/** Existing active notes; undefined when the slide has a null notes reference. */
	readonly notes: PptNotes | undefined;
	readonly notesStatus: 'present' | 'absent' | 'unsupported';
	/** A decode refusal is distinct from a null notes reference. Bytes remain opaque. */
	readonly notesDiagnostic: string | undefined;
	constructor(slideId: number, persistId: number, texts: readonly PptText[], shapes: readonly PptShape[] = [], notes?: PptNotes, notesDiagnostic?: string) {
		this.slideId = slideId; this.persistId = persistId;
		this.texts = Object.freeze([...texts]); this.shapes = Object.freeze([...shapes]); this.notes = notes;
		this.notesDiagnostic = notesDiagnostic; this.notesStatus = notesDiagnostic ? 'unsupported' : notes ? 'present' : 'absent'; Object.freeze(this);
	}
}

/** One existing inline notes text slot; body is a validated PT_NotesBody, not a
 * guessed shape index. Other placeholders/fields are readable but not writable. */
export class PptNoteText {
	readonly shapeId: number; readonly headerOffset: number;
	#read: () => PptNoteTextAtom; #write: (text: string) => void;
	constructor(atom: PptNoteTextAtom, read: () => PptNoteTextAtom, write: (text: string) => void) {
		this.shapeId = atom.shapeId; this.headerOffset = atom.headerOffset;
		this.#read = read; this.#write = write; Object.freeze(this);
	}
	get text(): string { return this.#read().text; }
	set text(value: string) { if (typeof value !== 'string') throw new UnsupportedOle2EditError('Notes replacement must be a string'); this.#write(value); }
	get role(): PptNoteTextAtom['role'] { return this.#read().role; }
	get placeholderId(): number | undefined { return this.#read().placeholderId; }
	get encoding(): PptNoteTextAtom['encoding'] { return this.#read().encoding; }
	get editRefusal(): string | undefined { return this.#read().editRefusal; }
	get runs(): never { throw new UnsupportedOle2EditError('PPT notes rich-text runs are preserved but not decoded'); }
}
export class PptNotes {
	readonly slideId: number; readonly slidePersistId: number;
	readonly notesId: number; readonly persistId: number;
	readonly texts: readonly PptNoteText[];
	constructor(record: PptSlideNotes, texts: readonly PptNoteText[]) {
		this.slideId = record.slideId; this.slidePersistId = record.slidePersistId;
		this.notesId = record.notesId; this.persistId = record.persistId;
		this.texts = Object.freeze([...texts]); Object.freeze(this);
	}
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
	/** Field-order-preserving values for undecoded large client anchors. */
	get rawAnchorValues(): readonly number[] | undefined { const values = this.#read().anchor?.rawValues; return values && Object.freeze([...values]); }
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
	#notesCache: { revision: number; notes: PptSlideNotes[] };
	#runsCache: { revision: number; records: PptTextRunRecord[] } | undefined;
	get slides(): readonly PptSlide[] { return this.#slides; }
	get unsupported(): readonly string[] { return this.#unsupported; }

	constructor(input: Uint8Array) {
		super(input);
		const shapeModel = readPptSlideShapes(this.getBytes());
		const model = readPptSlideTexts(this.getBytes());
		let notesModel: PptSlideNotes[] = [], notesDiagnostic: string | undefined;
		try { notesModel = readPptSlideNotes(this.getBytes()); }
		catch (error) { notesDiagnostic = error instanceof Error ? error.message : 'Notes decoding unsupported'; }
		this.#shapeCache = { revision: this.revision, slides: shapeModel };
		this.#notesCache = { revision: this.revision, notes: notesModel };
		this.#unsupported = Object.freeze(['formatting', 'group-transforms', 'pictures-content', 'masters', 'notes-creation-and-inherited-fields', 'animations', 'embedded-objects', 'unresolved-outline-shape-text', 'variable-length-text', ...(notesDiagnostic ? ['notes-decoding'] : [])]);
		this.#slides = Object.freeze(model.slides.map((slide, index) => {
			const texts = slide.texts.map((atom, textIndex) => new PptText({ slideId: slide.slideId, persistId: slide.persistId, textIndex }, atom,
				(node, value) => this.#edit(node, value), node => this.#readTextRuns(node), (node, run, value) => this.#editRun(node, run, value)));
			const shapes = shapeModel[index]!.shapes.map(record => new PptShape(slide, record, record.textIndexes.map(textIndex => texts[textIndex]!),
				() => this.#readShape(slide.slideId, slide.persistId, record), value => this.#editShape(slide.slideId, slide.persistId, record, value)));
			const record = notesModel.find(note => note.slideId === slide.slideId && note.slidePersistId === slide.persistId);
			const notes = record && new PptNotes(record, record.texts.map(atom => new PptNoteText(atom,
				() => this.#readNote(record, atom), value => this.#editNote(record, atom, value))));
			return new PptSlide(slide.slideId, slide.persistId, texts, shapes, notes, notesDiagnostic);
		}));
	}

	#readTextRuns(node: PptText): PptTextRunRecord {
		if (!this.#runsCache || this.#runsCache.revision !== this.revision) this.#runsCache = { revision: this.revision, records: readPptCharacterRuns(this.getBytes()) };
		const matches = this.#runsCache.records.filter(record => record.slideId === node.slideId && record.persistId === node.persistId && record.headerOffset === node.headerOffset);
		if (matches.length !== 1 || matches[0]!.text !== node.text) throw new UnsupportedOle2EditError('PPT text formatting identity no longer matches or is ambiguous');
		return matches[0]!;
	}
	#editRun(node: PptText, original: PptCharacterRunRecord, fontSizePoints: number | undefined): void {
		const revision = this.revision, current = this.#readTextRuns(node);
		const matches = current.status === 'decoded' ? current.runs.filter(run => run.runOffset === original.runOffset && run.styleHeaderOffset === original.styleHeaderOffset && run.start === original.start && run.end === original.end && run.rawCount === original.rawCount && run.mask === original.mask && run.fontSizeOffset === original.fontSizeOffset) : [];
		if (matches.length !== 1) throw new UnsupportedOle2EditError('PPT character run identity no longer matches or is ambiguous');
		const run = matches[0]!;
		const result = editPptCharacterFontSize(this.getBytes(), { slideId: node.slideId, persistId: node.persistId, headerOffset: node.headerOffset,
			runOffset: run.runOffset, styleHeaderOffset: run.styleHeaderOffset, start: run.start, end: run.end,
			expectedFontSizePoints: run.directFontSizePoints, fontSizePoints });
		if (this.revision !== revision) throw new UnsupportedOle2EditError('Document changed during character formatting edit');
		if (result.status === 'unsupported') throw new UnsupportedOle2EditError(result.reason);
		if (result.status === 'edited') this.commitBytes(result.bytes);
	}

	#readNote(original: PptSlideNotes, atom: PptNoteTextAtom): PptNoteTextAtom {
		if (this.#notesCache.revision !== this.revision) {
			try { this.#notesCache = { revision: this.revision, notes: readPptSlideNotes(this.getBytes()) }; }
			catch (error) { throw new UnsupportedOle2EditError(error instanceof Error ? error.message : 'Notes decoding unsupported'); }
		}
		const notes = this.#notesCache.notes.filter(note => note.slideId === original.slideId && note.slidePersistId === original.slidePersistId && note.notesId === original.notesId && note.persistId === original.persistId && note.headerOffset === original.headerOffset);
		const matches = notes.length === 1 ? notes[0]!.texts.filter(text => text.shapeId === atom.shapeId && text.headerOffset === atom.headerOffset && text.shapeHeaderOffset === atom.shapeHeaderOffset && text.encoding === atom.encoding && text.placeholderId === atom.placeholderId) : [];
		if (matches.length !== 1) throw new UnsupportedOle2EditError('PPT notes identity no longer matches or is ambiguous');
		return matches[0]!;
	}
	#editNote(note: PptSlideNotes, atom: PptNoteTextAtom, text: string): void {
		const revision = this.revision, current = this.#readNote(note, atom);
		const result = editPptNotesText(this.getBytes(), { slideId: note.slideId, slidePersistId: note.slidePersistId, notesId: note.notesId, persistId: note.persistId,
			shapeId: atom.shapeId, expectedHeaderOffset: atom.headerOffset, expectedText: current.text, text });
		if (this.revision !== revision) throw new UnsupportedOle2EditError('Document changed during notes edit');
		if (result.status === 'unsupported') throw new UnsupportedOle2EditError(result.reason);
		if (result.status === 'edited') this.commitBytes(result.bytes);
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
		if (!current.anchor?.bounds) throw new UnsupportedOle2EditError('PPT shape has no decoded anchor');
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


