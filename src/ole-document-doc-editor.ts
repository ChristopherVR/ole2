/* Adapted from ChristopherVR/pptx-viewer packages/core/src/core/utils, Apache-2.0. Original source: https://github.com/ChristopherVR/pptx-viewer. */
/**
 * In-place paragraph text editing for an embedded legacy binary Word
 * document (`Word.Document.8` / `.doc`, Word 97-2003) OLE payload.
 *
 * Companion to `ole-document-docx-editor.ts` (the modern `.docx` editor):
 * same public shape (`readOleDocParagraphs` / `writeOleDocParagraphEdit`,
 * `-Doc-` naming to keep the two apart), very different mechanics, because
 * `.doc` has no XML body to patch. Structures follow [MS-DOC], with fixture
 * regression tests and independent Word checks for the covered edit strategies
 * (see module docs on
 * `ole-document-doc-fib.ts`, `-pieces.ts`, `-fkp.ts`, `-cfb.ts`).
 *
 * ## Write strategy (why it is safe)
 *
 * Equal-length edits that fit each original piece's encoding patch only the
 * target text bytes, preserving every original run and CP-keyed feature table.
 * Structural target paragraphs (fields, objects, table/section marks) are refused.
 * Otherwise an edit APPENDS the new
 * paragraph text (plus a fresh paragraph mark) as a brand-new "piece" at the
 * end of the `WordDocument` stream, together with a brand-new single-run
 * PAPX/CHPX formatting page copied VERBATIM from the edited paragraph's own
 * original formatting (so alignment/spacing/font survive, mirroring the
 * `.docx` editor's "keep `w:pPr`, reuse the first run's `w:rPr`" scope), then
 * rewrites the piece table (and only the piece table: FKP bookkeeping tables
 * are extended, not rewritten) so the OLD range is simply unreferenced. This
 * is exactly the shape of a real Word "fast save" and was chosen because it
 * never has to shift or reinterpret bytes it does not fully understand.
 *
 * ## Scope limits (checked, not assumed)
 *
 * Character-count-changing edits require a document whose only populated character-position-keyed table
 * is the (mandatory, always present) section table `PlcfSed`, which this
 * module updates correctly for any number of sections. A document that also
 * uses footnotes, comments/annotations, fields, or bookmarks is left
 * UNCHANGED (`writeOleDocParagraphEdit` returns the original bytes; the
 * `tryWriteOleDocParagraphEdit` API returns an explicit rejection reason): those
 * features add their own character-position-keyed tables that a paragraph
 * edit's character-count shift would silently invalidate, and this module
 * has not been verified against a real Word round trip for that case. Non-main
 * stories and drawing/anchor tables are also guarded. The flat CFB rebuild
 * refuses nested storage containers for growing streams; CP-stable edits use
 * the hierarchy-preserving CFB patcher. Encrypted/unsupported FIBs are unreadable.
 * Reading (`readOleDocParagraphs`) otherwise has no editing feature restriction.
 *
 * @module ole-document-doc-editor
 */
import { unwrapDocBytes } from './ole-document-doc-cfb.js';
import { readDocFib, readFcLcbAt, patchDocFib } from './ole-document-doc-fib.js';
import type { DocFib, FcLcb } from './ole-document-doc-fib.js';
import {
	buildBteTableBytes,
	buildSingleRunChpxPage,
	buildSingleRunPapxPage,
	extractChpxBlob,
	extractPapxBlob,
	parseBteTable,
} from './ole-document-doc-fkp.js';
import {
	buildClxBytes,
	decodePiecesText,
	encodePieceText,
	parsePieceTable,
	parseDocClx,
	replacePieceRange,
} from './ole-document-doc-pieces.js';
import type { DocPiece } from './ole-document-doc-pieces.js';
import { encodeCp1252 } from './ole-document-doc-cp1252.js';

/** [MS-DOC] 2.5.3 `FibRgFcLcb97` indices for features this editor refuses to edit around (see module doc). */
const RISKY_PLCF_INDICES = [2, 3, 4, 5, 16, 17, 18, 19, 20, 40, 41, 42, 43, 46, 47, 48, 54, 55, 56, 57, 58, 59, 75, 76, 89, 90];
const RISKY_STTB_INDICES = [21, 22, 23]; // bkmk sttb, bkf, bkl
// [MS-DOC] 2.5.7/2.5.8/2.5.10: ignored table-character cache, email metadata,
// revision-save IDs, and ignored theme/color mapping. None defines live CP ranges.
const SAFE_EXTENDED_FC_LCB_INDICES = new Set([93, 94, 113, 181, 182]);

function hasUnsupportedFeatures(wordDoc: Uint8Array, fib: DocFib): boolean {
	if (fib.ccpOtherStories > 0) return true;
	// Effective versions can be extended even when FibBase.nFib remains Word97.
	// Refuse every populated unknown extension, including factoid/repair bookmarks.
	for (let index = 93; index < fib.fibRgFcLcbCount; index++) {
		if (!SAFE_EXTENDED_FC_LCB_INDICES.has(index) && readFcLcbAt(wordDoc, fib, index).lcb !== 0)
			return true;
	}
	// A trivial/empty PLCF is still (n+1)*4 = 4 bytes (a single sentinel CP);
	// anything larger means the document actually uses the feature.
	for (const index of RISKY_PLCF_INDICES) {
		if (index >= fib.fibRgFcLcbCount) continue;
		if (readFcLcbAt(wordDoc, fib, index).lcb > 4) {
			return true;
		}
	}
	for (const index of RISKY_STTB_INDICES) {
		if (readFcLcbAt(wordDoc, fib, index).lcb > 0) {
			return true;
		}
	}
	return false;
}

/** Locate the piece covering CP `cp`, and its byte (FC) offset within `WordDocument`. */
function pieceAndFcAtCp(pieces: readonly DocPiece[], cp: number): { piece: DocPiece; fc: number } {
	for (const piece of pieces) {
		if (cp >= piece.cpStart && cp < piece.cpEnd) {
			const unit = piece.compressed ? 1 : 2;
			return { piece, fc: piece.fc + (cp - piece.cpStart) * unit };
		}
	}
	throw new Error(`CP ${cp} is not covered by any piece`);
}

/** Paragraph boundaries (CP just past each paragraph mark) within the main body text `[0, ccpText)`. */
function paragraphBoundaries(bodyText: string): number[] {
	const boundaries: number[] = [];
	for (let i = 0; i < bodyText.length; i++) {
		if (bodyText[i] === '\r') {
			boundaries.push(i + 1);
		}
	}
	if (boundaries[boundaries.length - 1] !== bodyText.length && bodyText.length > 0) {
		boundaries.push(bodyText.length);
	}
	return boundaries;
}

function encodePieceTextUtf16(text: string): Uint8Array {
	const bytes = new Uint8Array(text.length * 2);
	const view = new DataView(bytes.buffer);
	for (let i = 0; i < text.length; i++) view.setUint16(i * 2, text.charCodeAt(i), true);
	return bytes;
}

/** Read every main-body paragraph's plain text (paragraph mark excluded) from an embedded `.doc` payload. */
export function readOleDocParagraphs(docBytes: Uint8Array): string[] | undefined {
	try {
		const cfb = unwrapDocBytes(docBytes);
		if (!cfb) {
			return undefined;
		}
		const fib = readDocFib(cfb.wordDocBytes);
		const pieces = parsePieceTable(cfb.tableBytes, fib.clx);
		const totalCp = pieces.at(-1)?.cpEnd ?? 0;
		if (totalCp > fib.ccpText + fib.ccpOtherStories + 1) return undefined;
		const fullText = decodePiecesText(cfb.wordDocBytes, pieces, fib.ccpText);
		if (fullText.length < fib.ccpText) return undefined;
		const bodyText = fullText.slice(0, fib.ccpText);
		const boundaries = paragraphBoundaries(bodyText);
		let start = 0;
		const paragraphs: string[] = [];
		for (const end of boundaries) {
			const raw = bodyText.slice(start, end);
			paragraphs.push(raw.endsWith('\r') ? raw.slice(0, -1) : raw);
			start = end;
		}
		return paragraphs;
	} catch {
		return undefined;
	}
}

/** Shift (or leave alone) every CP in a `PlcfSed`'s boundary array in place, for a same-size character-count-changing edit. */
function shiftSectionTableCps(
	tableBytes: Uint8Array,
	sed: FcLcb,
	editEndCp: number,
	delta: number,
): void {
	const view = new DataView(tableBytes.buffer, tableBytes.byteOffset, tableBytes.byteLength);
	if (sed.lcb < 4 || (sed.lcb - 4) % 16 !== 0 || sed.fc > tableBytes.length - sed.lcb)
		throw new Error('Invalid section table');
	const n = (sed.lcb - 4) / 16;
	for (let i = 0; i <= n; i++) {
		const off = sed.fc + i * 4;
		const cp = view.getInt32(off, true);
		if (cp >= editEndCp) {
			view.setInt32(off, cp + delta, true);
		}
	}
}

/**
 * Replace one main-body paragraph's text in an embedded `.doc` payload,
 * preserving that paragraph's own paragraph/character formatting (see module
 * doc). Returns the original bytes unchanged if the edit could not be safely
 * applied (unreadable payload, out-of-range paragraph index, or a document
 * feature this editor does not support around, per module doc).
 */
export function writeOleDocParagraphEdit(
	docBytes: Uint8Array,
	paragraphIndex: number,
	text: string,
): Uint8Array {
	return tryWriteOleDocParagraphEdit(docBytes, paragraphIndex, text).bytes;
}

/** Explicit edit outcome; rejected edits always retain the original input bytes. */
export type OleDocParagraphEditResult =
	| { status: 'edited'; bytes: Uint8Array; strategy: 'preserved-runs' | 'piece-append' }
	| { status: 'rejected'; bytes: Uint8Array; reason: 'invalid-payload' | 'unsupported-container' |
		'unsupported-features' | 'invalid-paragraph-index' | 'invalid-text' | 'invalid-document' };

/** Apply the same paragraph edit as the compatibility API, with an explicit rejection reason. */
export function tryWriteOleDocParagraphEdit(
	docBytes: Uint8Array,
	paragraphIndex: number,
	text: string,
): OleDocParagraphEditResult {
	const rejected = (reason: Extract<OleDocParagraphEditResult, { status: 'rejected' }>['reason']): OleDocParagraphEditResult =>
		({ status: 'rejected', bytes: docBytes, reason });
	if (!Number.isInteger(paragraphIndex) || paragraphIndex < 0) return rejected('invalid-paragraph-index');
	if (/[\u0000-\u0008\u000b\u000c\u000e-\u001f]/u.test(text)) return rejected('invalid-text');
	try {
		const cfb = unwrapDocBytes(docBytes);
		if (!cfb) {
			return rejected('invalid-payload');
		}
		const fib = readDocFib(cfb.wordDocBytes);
		const { pieces, prcBytes } = parseDocClx(cfb.tableBytes, fib.clx);
		const totalCp = pieces.at(-1)?.cpEnd ?? 0;
		if (totalCp > fib.ccpText + fib.ccpOtherStories + 1) return rejected('invalid-document');
		const fullText = decodePiecesText(cfb.wordDocBytes, pieces, fib.ccpText);
		if (fullText.length < fib.ccpText) return rejected('invalid-document');
		const bodyText = fullText.slice(0, fib.ccpText);
		const boundaries = paragraphBoundaries(bodyText);
		const paragraphEndCp = boundaries[paragraphIndex];
		if (paragraphEndCp === undefined) {
			return rejected('invalid-paragraph-index');
		}
		const paragraphStartCp = paragraphIndex === 0 ? 0 : boundaries[paragraphIndex - 1]!;
		if (/[\u0000-\u0008\u000b\u000c\u000e-\u001f]/u.test(bodyText.slice(paragraphStartCp, paragraphEndCp)))
			return rejected('unsupported-features');
		const sanitizedText = text.replaceAll(/[\r\n]+/gu, ' ');
		// A CP-stable replacement can retain every piece, formatting run and CP-keyed
		// feature table, including features outside this paragraph. Only plain target
		// paragraphs qualify, and compressed pieces must retain their encoding.
		if (sanitizedText.length + 1 === paragraphEndCp - paragraphStartCp && bodyText[paragraphEndCp - 1] === '\r') {
			const replacement = `${sanitizedText}\r`;
			const patched = cfb.wordDocBytes.slice();
			let canPatch = true;
			for (const piece of pieces) {
				const start = Math.max(paragraphStartCp, piece.cpStart);
				const end = Math.min(paragraphEndCp, piece.cpEnd);
				if (start >= end) continue;
				const part = replacement.slice(start - paragraphStartCp, end - paragraphStartCp);
				const unit = piece.compressed ? 1 : 2;
				const fc = piece.fc + (start - piece.cpStart) * unit;
				const fcEnd = fc + part.length * unit;
				// Fast-save pieces can alias physical text; do not alter another logical range.
				if (pieces.some((other) => other !== piece && other.fc < fcEnd &&
					other.fc + (other.cpEnd - other.cpStart) * (other.compressed ? 1 : 2) > fc)) {
					canPatch = false; break;
				}
				const encoded = piece.compressed ? encodeCp1252(part) : encodePieceTextUtf16(part);
				if (!encoded) { canPatch = false; break; }
				patched.set(encoded, fc);
			}
			if (canPatch) return { status: 'edited', strategy: 'preserved-runs', bytes: cfb.rewrap(patched, cfb.tableBytes) };
		}
		if (!cfb.canRewrite) return rejected('unsupported-container');
		if (hasUnsupportedFeatures(cfb.wordDocBytes, fib)) return rejected('unsupported-features');

		const { piece: startPiece, fc: startFc } = pieceAndFcAtCp(pieces, paragraphStartCp);
		if (pieces.some((piece) => piece.cpStart < paragraphEndCp && piece.cpEnd > paragraphStartCp && piece.prm !== startPiece.prm))
			return rejected('unsupported-features');
		const { fc: paragraphMarkFc } = pieceAndFcAtCp(pieces, paragraphEndCp - 1);
		const papxBte = parseBteTable(cfb.tableBytes, fib.plcfbtePapx);
		const chpxBte = parseBteTable(cfb.tableBytes, fib.plcfbteChpx);
		const papxBlob = extractPapxBlob(cfb.wordDocBytes, papxBte, paragraphMarkFc);
		const chpxBlob = extractChpxBlob(cfb.wordDocBytes, chpxBte, startFc);

		const { bytes: textBytes, compressed } = encodePieceText(`${sanitizedText}\r`);

		const originalLen = cfb.wordDocBytes.length;
		const newTextFc = originalLen;
		const afterText = originalLen + textBytes.length;
		const padLen = (512 - (afterText % 512)) % 512;
		const papxPageOffset = afterText + padLen;
		const chpxPageOffset = papxPageOffset + 512;
		const newWordDocLen = chpxPageOffset + 512;

		const newWordDoc = new Uint8Array(newWordDocLen);
		newWordDoc.set(cfb.wordDocBytes, 0);
		newWordDoc.set(textBytes, originalLen);
		newWordDoc.set(buildSingleRunPapxPage(newTextFc, afterText, papxBlob), papxPageOffset);
		newWordDoc.set(buildSingleRunChpxPage(newTextFc, afterText, chpxBlob), chpxPageOffset);

		const newPapxPageNumber = papxPageOffset / 512;
		const newChpxPageNumber = chpxPageOffset / 512;
		// Extend the LAST existing range's boundary forward to `newTextFc`
		// (rather than inserting a separate filler range for the dead space
		// between the old text and the newly appended piece): a real
		// Word-COM round trip rejected the 3-range/filler-page form as
		// corrupt, while stretching the existing last range to reuse its own
		// page number for the now-larger, still just-two-ranges-longer table
		// opens cleanly. Nothing ever looks up an FC in the dead zone (no
		// piece maps there), so which page nominally "covers" it is moot;
		// what matters is which SHAPE of BTE table Word accepts.
		const newPapxBte = {
			fcs: [...papxBte.fcs.slice(0, -1), newTextFc, afterText],
			pns: [...papxBte.pns, newPapxPageNumber],
		};
		const newChpxBte = {
			fcs: [...chpxBte.fcs.slice(0, -1), newTextFc, afterText],
			pns: [...chpxBte.pns, newChpxPageNumber],
		};

		const newPieces = replacePieceRange(pieces, paragraphStartCp, paragraphEndCp, {
			fc: newTextFc,
			compressed,
			charLength: sanitizedText.length + 1,
			flagsWord: startPiece.flagsWord,
			prm: startPiece.prm,
		});
		const delta = sanitizedText.length + 1 - (paragraphEndCp - paragraphStartCp);
		const newCcpText = fib.ccpText + delta;

		const clxBytes = buildClxBytes(newPieces, prcBytes);
		const papxBteBytes = buildBteTableBytes(newPapxBte);
		const chpxBteBytes = buildBteTableBytes(newChpxBte);

		const newTableBytes = new Uint8Array(
			cfb.tableBytes.length + clxBytes.length + papxBteBytes.length + chpxBteBytes.length,
		);
		newTableBytes.set(cfb.tableBytes, 0);
		let tableOff = cfb.tableBytes.length;
		const newClx: FcLcb = { fc: tableOff, lcb: clxBytes.length };
		newTableBytes.set(clxBytes, tableOff);
		tableOff += clxBytes.length;
		const newPapxBteFcLcb: FcLcb = { fc: tableOff, lcb: papxBteBytes.length };
		newTableBytes.set(papxBteBytes, tableOff);
		tableOff += papxBteBytes.length;
		const newChpxBteFcLcb: FcLcb = { fc: tableOff, lcb: chpxBteBytes.length };
		newTableBytes.set(chpxBteBytes, tableOff);

		shiftSectionTableCps(newTableBytes, fib.sed, paragraphEndCp, delta);

		patchDocFib(newWordDoc, fib, {
			ccpText: newCcpText,
			cbMac: newWordDocLen,
			clx: newClx,
			plcfbteChpx: newChpxBteFcLcb,
			plcfbtePapx: newPapxBteFcLcb,
		});

		return { status: 'edited', strategy: 'piece-append', bytes: cfb.rewrap(newWordDoc, newTableBytes) };
	} catch {
		return rejected('invalid-document');
	}
}
