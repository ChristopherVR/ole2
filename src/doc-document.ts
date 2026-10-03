import { Ole2DocumentBase, Ole2DocumentError, UnsupportedOle2EditError } from './ole2-document-base.js';
import { readOleDocParagraphs, tryWriteOleDocParagraphEdit } from './ole-document-doc-editor.js';
import { readDocCharacterRuns, readDocParagraphStyleIndices, writeDocCharacterRunFlag } from './ole-document-doc-runs.js';
import type { DocSprm } from './ole-document-doc-runs.js';
import type { ParsedDocCharacterRun } from './ole-document-doc-runs.js';

/** Direct CHPX exceptions. Undefined flags are inherited or undecoded, rather
 * than false. Handles are snapshots and refuse mutation after any other edit. */
export interface DocCharacterRun {
	readonly cpStart: number;
	readonly cpEnd: number;
	readonly text: string;
	directBold: boolean | undefined;
	directItalic: boolean | undefined;
	readonly directFontSizePoints: number | undefined;
	readonly sprms: readonly DocSprm[];
}

/** Plain main-story paragraph text. Structural content remains readable, but
 * the setter refuses edits that the preservation-safe DOC writer cannot apply. */
export interface DocParagraph {
	readonly index: number;
	text: string;
	readonly runs?: readonly DocCharacterRun[];
	readonly styleIndex?: number | undefined;
}

/** Rich getters are present on paragraphs returned by DocDocument. Optional
 * additions on DocParagraph retain compatibility with caller-created values. */
export interface ParsedDocParagraph extends DocParagraph {
	readonly runs: readonly DocCharacterRun[];
	readonly styleIndex: number | undefined;
}

const DOC_CAPABILITIES = Object.freeze({
	read: Object.freeze(['paragraph-text', 'character-runs', 'direct-character-formatting', 'paragraph-style-index', 'compound-streams']),
	write: Object.freeze(['paragraph-text', 'existing-direct-bold-italic']),
	limitations: Object.freeze([
		'Paragraph insertion, removal and embedded paragraph breaks are unsupported.',
		'Character formatting reports direct CHPX exceptions, without resolving styles or piece PRMs.',
		'Bold/italic setters force an absolute value in an existing exclusive understood CHPX run. Run handles expire after edits.',
		'Fields, tables, objects and inherited styles are preserved where supported; their models are not editable.',
		'Unsupported edits and processing limits throw without changing the document.',
		'Text uses the bounded DOC codec defaults: 16,777,216 main-story characters and 65,536 pieces or field records.',
		'Rich formatting is bounded to 65,536 physical/mapped runs and paragraph style records.',
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
				get runs() { return owner.#paragraphRuns(index); },
				get styleIndex() {
					try { return owner.#paragraphStyles()[index]; }
					catch { throw new UnsupportedOle2EditError('unsupported-formatting'); }
				},
			});
		}));
	}

	get paragraphs(): readonly ParsedDocParagraph[] { return this.#paragraphs; }

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
				let bold = run.directBold, italic = run.directItalic, sprms = run.sprms;
				const set = (flag: 'bold' | 'italic', value: boolean | undefined) => {
					if (owner.revision !== revision) throw new UnsupportedOle2EditError('stale-run');
					if (typeof value !== 'boolean') throw new UnsupportedOle2EditError('invalid-formatting');
					let bytes: Uint8Array;
					try { bytes = writeDocCharacterRunFlag(owner.getBytes(), cpStart, cpEnd, flag, value); }
					catch (error) { throw new UnsupportedOle2EditError(error instanceof Error ? error.message : 'unsupported-formatting'); }
					const text = readOleDocParagraphs(bytes);
					if (!text || text.length !== owner.#text.length || text.some((t, i) => t !== owner.#text[i])) throw new UnsupportedOle2EditError('invalid-document');
					readDocCharacterRuns(bytes); // Revalidate the candidate before the atomic commit.
					owner.commitBytes(bytes);
					if (flag === 'bold') bold = value; else italic = value;
					sprms = Object.freeze(sprms.map((p) => p.opcode === (flag === 'bold' ? 0x0835 : 0x0836) ?
						Object.freeze({ opcode: p.opcode, operand: Object.freeze([Number(value)]) }) : p));
				};
				return Object.freeze({ cpStart, cpEnd, text: run.text.slice(cpStart - run.cpStart, cpEnd - run.cpStart), directFontSizePoints: run.directFontSizePoints, get sprms() { return sprms; },
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
