import { Ole2DocumentBase, Ole2DocumentError, UnsupportedOle2EditError } from './ole2-document-base.js';
import { readOleDocParagraphs, tryWriteOleDocParagraphEdit } from './ole-document-doc-editor.js';
import { readDocCharacterRuns, readDocParagraphStyleIndices, writeDocCharacterRunFlag, writeDocCharacterRunFontSize } from './ole-document-doc-runs.js';
import type { DocSprm } from './ole-document-doc-runs.js';
import type { ParsedDocCharacterRun } from './ole-document-doc-runs.js';
import { readDocParagraphAlignments, writeDocParagraphAlignment } from './ole-document-doc-paragraph-format.js';

/** Logical alignment relative to paragraph direction; not resolved physical layout. */
export type DocParagraphAlignment = 'start' | 'center' | 'end' | 'justify';

/** Direct CHPX exceptions. Undefined flags are inherited or undecoded, rather
 * than false. Handles are snapshots and refuse mutation after any other edit. */
export interface DocCharacterRun {
	readonly cpStart: number;
	readonly cpEnd: number;
	readonly text: string;
	directBold: boolean | undefined;
	directItalic: boolean | undefined;
	directFontSizePoints: number | undefined;
	readonly sprms: readonly DocSprm[];
}

/** Plain main-story paragraph text. Structural content remains readable, but
 * the setter refuses edits that the preservation-safe DOC writer cannot apply. */
export interface DocParagraph {
	readonly index: number;
	text: string;
	readonly runs?: readonly DocCharacterRun[];
	readonly styleIndex?: number | undefined;
	directAlignment?: DocParagraphAlignment | undefined;
}

/** Rich getters are present on paragraphs returned by DocDocument. Optional
 * additions on DocParagraph retain compatibility with caller-created values. */
export interface ParsedDocParagraph extends DocParagraph {
	readonly runs: readonly DocCharacterRun[];
	readonly styleIndex: number | undefined;
	directAlignment: DocParagraphAlignment | undefined;
}

const DOC_CAPABILITIES = Object.freeze({
	read: Object.freeze(['paragraph-text', 'character-runs', 'direct-character-formatting', 'paragraph-style-index', 'direct-paragraph-alignment', 'compound-streams']),
	write: Object.freeze(['paragraph-text', 'existing-direct-bold-italic', 'existing-direct-font-size', 'existing-direct-paragraph-alignment']),
	limitations: Object.freeze([
		'Paragraph insertion, removal and embedded paragraph breaks are unsupported.',
		'Character formatting reports direct CHPX exceptions, without resolving styles or piece PRMs.',
		'Bold/italic setters force an absolute value in an existing exclusive understood CHPX run. Run handles expire after edits.',
		'Font-size setters replace an existing exclusive sprmCHps slot with 1..1638 points in exact half-point increments.',
		'Fields, tables, objects and inherited styles are preserved where supported; their models are not editable.',
		'Unsupported edits and processing limits throw without changing the document.',
		'Text uses the bounded DOC codec defaults: 16,777,216 main-story characters and 65,536 pieces or field records.',
		'Rich formatting is bounded to 65,536 physical/mapped runs and paragraph style records.',
		'Logical paragraph alignment replaces a sole existing exclusive sprmPJc slot. Matching legacy mirrors support center/justify only, with both slots updated. Styles, legacy-only alignment, tables and opaque formatting refuse.',
		'Paragraph alignment inspection is bounded to 262,144 SPRM records before allocation, including shared or stale formatting runs.',
	]),
});

/** Editable legacy Word document. The paragraph collection has fixed structure;
 * text setters preserve supported formatting and commit only validated results. */
export class DocDocument extends Ole2DocumentBase {
	get kind(): 'doc' { return 'doc'; }
	get capabilities() { return DOC_CAPABILITIES; }
	readonly #paragraphs: readonly ParsedDocParagraph[];
	#text: readonly string[];
	#runsCache: { revision: number; runs: readonly ParsedDocCharacterRun[] } | undefined;
	#stylesCache: { revision: number; styles: readonly (number | undefined)[] } | undefined;
	#alignmentsCache: { revision: number; alignments: readonly (DocParagraphAlignment | undefined)[] } | undefined;

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
				get directAlignment() {
					try { return owner.#paragraphAlignments()[index]; }
					catch { throw new UnsupportedOle2EditError('unsupported-formatting'); }
				},
				set directAlignment(value: DocParagraphAlignment | undefined) { owner.#setParagraphAlignment(index, value); },
				get runs() { return owner.#paragraphRuns(index); },
				get styleIndex() {
					try { return owner.#paragraphStyles()[index]; }
					catch { throw new UnsupportedOle2EditError('unsupported-formatting'); }
				},
			});
		}));
	}

	get paragraphs(): readonly ParsedDocParagraph[] { return this.#paragraphs; }

	#paragraphAlignments(): readonly (DocParagraphAlignment | undefined)[] {
		if (this.#alignmentsCache?.revision !== this.revision)
			this.#alignmentsCache = { revision: this.revision, alignments: readDocParagraphAlignments(this.getBytes()) };
		return this.#alignmentsCache.alignments;
	}

	#setParagraphAlignment(index: number, value: DocParagraphAlignment | undefined): void {
		if (typeof value !== 'string') throw new UnsupportedOle2EditError('invalid-formatting');
		let bytes: Uint8Array;
		try { bytes = writeDocParagraphAlignment(this.getBytes(), index, value); }
		catch (error) { throw new UnsupportedOle2EditError(error instanceof Error ? error.message : 'unsupported-formatting'); }
		const text = readOleDocParagraphs(bytes), alignments = readDocParagraphAlignments(bytes);
		if (!text || text.length !== this.#text.length || text.some((t, i) => t !== this.#text[i]) || alignments[index] !== value)
			throw new UnsupportedOle2EditError('invalid-document');
		this.commitBytes(bytes);
	}

	#currentRuns(): readonly ParsedDocCharacterRun[] {
		if (this.#runsCache?.revision !== this.revision)
			this.#runsCache = { revision: this.revision, runs: readDocCharacterRuns(this.getBytes()) };
		return this.#runsCache.runs;
	}

	#paragraphStyles(): readonly (number | undefined)[] {
		if (this.#stylesCache?.revision !== this.revision)
			this.#stylesCache = { revision: this.revision, styles: readDocParagraphStyleIndices(this.getBytes()) };
		return this.#stylesCache.styles;
	}

	#paragraphRuns(index: number): readonly DocCharacterRun[] {
		let start = 0;
		for (let i = 0; i < index; i++) start += this.#text[i]!.length + 1;
		const end = start + this.#text[index]!.length, owner = this, revision = this.revision;
		try {
			return Object.freeze(this.#currentRuns().filter((r) => r.cpStart < end && r.cpEnd > start).map((run) => {
				const cpStart = Math.max(start, run.cpStart), cpEnd = Math.min(end, run.cpEnd);
				let bold = run.directBold, italic = run.directItalic, size = run.directFontSizePoints, sprms = run.sprms;
				const set = (flag: 'bold' | 'italic' | 'size', value: boolean | number | undefined) => {
					if (owner.revision !== revision) throw new UnsupportedOle2EditError('stale-run');
					if (flag === 'size' ? typeof value !== 'number' : typeof value !== 'boolean') throw new UnsupportedOle2EditError('invalid-formatting');
					let bytes: Uint8Array;
					try { bytes = flag === 'size' ? writeDocCharacterRunFontSize(owner.getBytes(), cpStart, cpEnd, value as number) : writeDocCharacterRunFlag(owner.getBytes(), cpStart, cpEnd, flag, value as boolean); }
					catch (error) { throw new UnsupportedOle2EditError(error instanceof Error ? error.message : 'unsupported-formatting'); }
					const text = readOleDocParagraphs(bytes);
					if (!text || text.length !== owner.#text.length || text.some((t, i) => t !== owner.#text[i])) throw new UnsupportedOle2EditError('invalid-document');
					readDocCharacterRuns(bytes); // Revalidate the candidate before the atomic commit.
					owner.commitBytes(bytes);
					if (flag === 'bold') bold = value as boolean; else if (flag === 'italic') italic = value as boolean; else size = value as number;
					const opcode = flag === 'bold' ? 0x0835 : flag === 'italic' ? 0x0836 : 0x4a43;
					const operand = flag === 'size' ? [(value as number) * 2 & 255, (value as number) * 2 >>> 8] : [Number(value)];
					sprms = Object.freeze(sprms.map((p) => p.opcode === opcode ?
						Object.freeze({ opcode: p.opcode, operand: Object.freeze(operand) }) : p));
				};
				return Object.freeze({ cpStart, cpEnd, text: run.text.slice(cpStart - run.cpStart, cpEnd - run.cpStart), get directFontSizePoints() { return size; }, set directFontSizePoints(value: number | undefined) { set('size', value); }, get sprms() { return sprms; },
					get directBold() { return bold; }, set directBold(value: boolean | undefined) { set('bold', value); },
					get directItalic() { return italic; }, set directItalic(value: boolean | undefined) { set('italic', value); },
				});
			}));
		} catch (error) {
			if (error instanceof UnsupportedOle2EditError) throw error;
			throw new UnsupportedOle2EditError('unsupported-formatting');
		}
	}

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
