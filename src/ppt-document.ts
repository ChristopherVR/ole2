/** Active binary PPT text model backed by preservation-safe stream edits. */
import { Ole2DocumentBase, UnsupportedOle2EditError } from './ole2-document-base.js';
import { editPptSlideText, readPptSlideTexts, type PptTextAtom } from './legacy-ppt-text.js';

const PPT_CAPABILITIES = Object.freeze({
	read: Object.freeze(['active-slide-ids', 'active-outline-and-inline-text', 'text-encoding', 'opaque-streams']),
	write: Object.freeze(['existing-text-fixed-utf16-length']),
	limitations: Object.freeze(['shapes-and-shape-to-text-references', 'rich-text-runs', 'notes', 'variable-length-text', 'OOXML-mirrored-text', 'geometry', 'animations', 'masters', 'embedded-objects']),
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
	/** Active outline atoms followed by inline atoms; shape refs are unresolved. */
	readonly texts: readonly PptText[];
	constructor(slideId: number, persistId: number, texts: readonly PptText[]) {
		this.slideId = slideId; this.persistId = persistId;
		this.texts = Object.freeze([...texts]); Object.freeze(this);
	}
	get shapes(): never { throw new UnsupportedOle2EditError('PPT shapes and shape-to-text references are not decoded by this adapter'); }
	get notes(): never { throw new UnsupportedOle2EditError('PPT notes are preserved but are not decoded by this adapter'); }
}

/** Indexed active slides with transactional fixed-length text edits. Unsupported
 * geometry, formatting, notes, mirrors and unknown records remain in the CFB.
 * This model does not claim that text atoms correspond one-to-one to shapes. */
export class PptDocument extends Ole2DocumentBase {
	get kind(): 'ppt' { return 'ppt'; }
	get capabilities(): typeof PPT_CAPABILITIES { return PPT_CAPABILITIES; }
	readonly #slides: readonly PptSlide[];
	readonly #unsupported: readonly string[];
	get slides(): readonly PptSlide[] { return this.#slides; }
	get unsupported(): readonly string[] { return this.#unsupported; }

	constructor(input: Uint8Array) {
		super(input);
		const model = readPptSlideTexts(this.getBytes());
		this.#unsupported = Object.freeze([...model.unsupported, 'shape-to-text-references', 'variable-length-text']);
		this.#slides = Object.freeze(model.slides.map(slide => new PptSlide(slide.slideId, slide.persistId,
			slide.texts.map((atom, textIndex) => new PptText({ slideId: slide.slideId, persistId: slide.persistId, textIndex }, atom,
				(node, value) => this.#edit(node, value))))));
	}

	#edit(node: PptText, value: string): void {
		const bytes = this.getBytes(), current = readPptSlideTexts(bytes);
		const slideIndex = current.slides.findIndex(slide => slide.slideId === node.slideId && slide.persistId === node.persistId);
		const atom = current.slides[slideIndex]?.texts[node.textIndex];
		if (!atom || atom.headerOffset !== node.headerOffset || atom.encoding !== node.encoding || atom.text !== node.text)
			throw new UnsupportedOle2EditError('PPT text identity or expected text no longer matches');
		const result = editPptSlideText(bytes, { slideIndex, textIndex: node.textIndex, expectedText: node.text, text: value });
		if (result.status === 'unsupported') throw new UnsupportedOle2EditError(result.reason);
		if (result.status === 'edited') this.commitBytes(result.bytes);
	}
}
