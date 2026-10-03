import { Ole2DocumentBase, Ole2DocumentError, UnsupportedOle2EditError } from './ole2-document-base.js';
import { readOleDocParagraphs, tryWriteOleDocParagraphEdit } from './ole-document-doc-editor.js';

/** Plain main-story paragraph text. Structural content remains readable, but
 * the setter refuses edits that the preservation-safe DOC writer cannot apply. */
export interface DocParagraph {
	readonly index: number;
	text: string;
}

const DOC_CAPABILITIES = Object.freeze({
	read: Object.freeze(['paragraph-text', 'compound-streams']),
	write: Object.freeze(['paragraph-text']),
	limitations: Object.freeze([
		'Paragraph insertion, removal and embedded paragraph breaks are unsupported.',
		'Formatting, fields, tables and objects are preserved where supported; their models are not editable.',
		'Unsupported edits and processing limits throw without changing the document.',
		'Text uses the bounded DOC codec defaults: 16,777,216 main-story characters and 65,536 pieces or field records.',
	]),
});

/** Editable legacy Word document. The paragraph collection has fixed structure;
 * text setters preserve supported formatting and commit only validated results. */
export class DocDocument extends Ole2DocumentBase {
	get kind(): 'doc' { return 'doc'; }
	get capabilities() { return DOC_CAPABILITIES; }
	readonly #paragraphs: readonly DocParagraph[];
	#text: readonly string[];

	constructor(input: Uint8Array) {
		super(input);
		const text = readOleDocParagraphs(this.getBytes());
		if (!text) throw new Ole2DocumentError('invalid-format', 'Not a supported unencrypted DOC document');
		this.#text = Object.freeze(text);
		this.#paragraphs = Object.freeze(text.map((_, index) => {
			const owner = this;
			return Object.freeze({
				index,
				get text() { return owner.#text[index]!; },
				set text(value: string) { owner.#setParagraphText(index, value); },
			});
		}));
	}

	get paragraphs(): readonly DocParagraph[] { return this.#paragraphs; }

	#setParagraphText(index: number, value: string): void {
		if (typeof value !== 'string') throw new UnsupportedOle2EditError('invalid-text');
		if (/[\r\n]/u.test(value)) throw new UnsupportedOle2EditError('paragraph-structure');
		if (value === this.#text[index]) return;
		const result = tryWriteOleDocParagraphEdit(this.getBytes(), index, value);
		if (result.status !== 'edited') throw new UnsupportedOle2EditError(result.reason);
		const text = readOleDocParagraphs(result.bytes);
		if (!text || text.length !== this.#text.length || text[index] !== value)
			throw new UnsupportedOle2EditError('invalid-document');
		this.commitBytes(result.bytes);
		this.#text = Object.freeze(text);
	}
}
