/**
 * Read-only, best-effort slide text recovery for damaged MS-PPT files.
 *
 * `readPptSlideTexts` refuses anything it cannot attribute exactly, which is
 * right for edits but leaves a damaged deck with no text at all. This
 * extractor degrades in steps and says which one it used:
 *
 * 1. `active`: the exact `readPptSlideTexts` model.
 * 2. `persist-scan`: the newest UserEditAtom found by scanning the stream
 *    (for a missing or wrong Current User offset) and its persist
 *    directory, tolerating malformed records slide by slide.
 * 3. `record-scan`: every top-level slide container in stream order, the
 *    approach of SheetJS `ppt`/`ppt-to-text`. Incremental saves can leave
 *    superseded slides in the stream, so this may repeat or resurrect text.
 *
 * Results are never valid edit targets. Encrypted documents are refused.
 *
 * @module legacy-ppt-text-salvage
 */

import { parseOle2 } from './ole2-parser-read.js';
import { readPptSlideTexts, PptTextError } from './legacy-ppt-text.js';
import { buildPersistDirectory } from './ppt/persist-directory.js';
import { RT, OA, HEADER_TOKEN_ENCRYPTED } from './legacy-ppt-record-types.js';

export type PptSalvageMode = 'active' | 'persist-scan' | 'record-scan';
export interface PptSalvagedText {
	mode: PptSalvageMode;
	/** One array of text blocks per slide; `\r` and vertical-tab breaks become `\n`. */
	slides: string[][];
	/** Why stricter modes were skipped, and caveats for the mode used. */
	diagnostics: string[];
}

interface Rec { type: number; ver: number; data: number; end: number }
const MAX_DEPTH = 64;

function recordAt(view: DataView, at: number, end: number): Rec | undefined {
	if (end - at < 8) return undefined;
	const verInst = view.getUint16(at, true);
	const len = view.getUint32(at + 4, true);
	if (len > end - at - 8) return undefined;
	return { type: view.getUint16(at + 2, true), ver: verInst & 0xf, data: at + 8, end: at + 8 + len };
}
/** Sibling records until the first one that does not fit; never throws. */
function records(view: DataView, start: number, end: number): Rec[] {
	const out: Rec[] = [];
	for (let at = start; at < end;) {
		const rec = recordAt(view, at, end);
		if (!rec) break;
		out.push(rec);
		at = rec.end;
	}
	return out;
}
const tidy = (text: string) => text.replace(/\r\n?|\v/g, '\n');
function atomText(view: DataView, rec: Rec): string | undefined {
	let text = '';
	if (rec.type === RT.TextCharsAtom) for (let at = rec.data; at + 1 < rec.end; at += 2) text += String.fromCharCode(view.getUint16(at, true));
	// TextBytesAtom holds low bytes of UTF-16 code units ([MS-PPT] 2.9.43).
	else if (rec.type === RT.TextBytesAtom) for (let at = rec.data; at < rec.end; at++) text += String.fromCharCode(view.getUint8(at));
	else return undefined;
	return tidy(text);
}
function containerText(view: DataView, container: Rec, depth = 0): string[] {
	if (depth > MAX_DEPTH) return [];
	const out: string[] = [];
	for (const rec of records(view, container.data, container.end)) {
		const text = atomText(view, rec);
		if (text !== undefined) out.push(text);
		else if (rec.ver === 0xf || rec.type === OA.ClientTextbox) out.push(...containerText(view, rec, depth + 1));
	}
	return out;
}

function persistScan(view: DataView, currentUserOffset: number | undefined, diagnostics: string[]): string[][] | undefined {
	const candidates = records(view, 0, view.byteLength).filter((rec) => rec.type === RT.UserEditAtom).map((rec) => rec.data - 8);
	if (currentUserOffset !== undefined && !candidates.includes(currentUserOffset)) candidates.push(currentUserOffset);
	candidates.sort((a, b) => a - b);
	for (let i = candidates.length - 1; i >= 0; i--) {
		let chain;
		try {
			chain = buildPersistDirectory(view, candidates[i]!);
		} catch (error) {
			diagnostics.push(`UserEditAtom at ${candidates[i]}: ${(error as Error).message}`);
			continue;
		}
		if (chain.currentEdit.encryptSessionPersistIdRef !== undefined) throw new PptTextError('encrypted', 'Encrypted PPT text is unsupported');
		const docOffset = chain.directory.get(chain.currentEdit.docPersistIdRef);
		const doc = docOffset === undefined ? undefined : recordAt(view, docOffset, view.byteLength);
		if (!doc || doc.type !== RT.Document) {
			diagnostics.push(`UserEditAtom at ${candidates[i]}: missing document container`);
			continue;
		}
		const slides: string[][] = [];
		for (const list of records(view, doc.data, doc.end)) {
			if (list.type !== RT.SlideListWithText || (view.getUint16(list.data - 8, true) >> 4) !== 0) continue;
			let current: { persistId: number; texts: string[] } | undefined;
			const flush = () => {
				if (!current) return;
				const offset = chain.directory.get(current.persistId);
				const slide = offset === undefined ? undefined : recordAt(view, offset, view.byteLength);
				if (slide?.type === RT.Slide) current.texts.push(...containerText(view, slide));
				else diagnostics.push(`Slide persist ${current.persistId}: missing slide container; outline text only`);
				slides.push(current.texts);
			};
			for (const rec of records(view, list.data, list.end)) {
				if (rec.type === RT.SlidePersistAtom && rec.end - rec.data >= 4) {
					flush();
					current = { persistId: view.getUint32(rec.data, true), texts: [] };
				} else {
					const text = atomText(view, rec);
					if (text !== undefined && current) current.texts.push(text);
				}
			}
			flush();
		}
		return slides;
	}
	return undefined;
}

/** Recover as much slide text as the file allows, preferring exact models. */
export function salvagePptText(input: Uint8Array): PptSalvagedText {
	const diagnostics: string[] = [];
	try {
		const model = readPptSlideTexts(input);
		return { mode: 'active', slides: model.slides.map((slide) => slide.texts.map((atom) => tidy(atom.text))), diagnostics };
	} catch (error) {
		if (error instanceof PptTextError && error.code === 'encrypted') throw error;
		diagnostics.push(`active: ${(error as Error).message}`);
	}
	const file = parseOle2(input);
	const stream = file.getStream('PowerPoint Document');
	if (!stream) throw new PptTextError('corrupt', 'Missing PowerPoint Document stream');
	if (file.getStream('EncryptedSummary')) throw new PptTextError('encrypted', 'Encrypted PPT text is unsupported');
	const view = new DataView(stream.buffer, stream.byteOffset, stream.byteLength);

	let currentUserOffset: number | undefined;
	try {
		const current = file.getStream('Current User');
		if (current && current.length >= 24) {
			const cv = new DataView(current.buffer, current.byteOffset, current.byteLength);
			if (cv.getUint32(12, true) === HEADER_TOKEN_ENCRYPTED) throw new PptTextError('encrypted', 'Encrypted PPT text is unsupported');
			currentUserOffset = cv.getUint32(16, true);
		}
	} catch (error) {
		if (error instanceof PptTextError) throw error;
		diagnostics.push(`Current User: ${(error as Error).message}`);
	}

	const scanned = persistScan(view, currentUserOffset, diagnostics);
	if (scanned?.length) return { mode: 'persist-scan', slides: scanned, diagnostics };

	const slides = records(view, 0, view.byteLength)
		.filter((rec) => rec.type === RT.Slide)
		.map((rec) => containerText(view, rec));
	diagnostics.push('record-scan: slides are in stream order and may include superseded saves');
	return { mode: 'record-scan', slides, diagnostics };
}
