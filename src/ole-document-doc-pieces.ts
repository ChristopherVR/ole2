/* Adapted from ChristopherVR/pptx-viewer packages/core/src/core/utils, Apache-2.0. Original source: https://github.com/ChristopherVR/pptx-viewer. */
/**
 * [MS-DOC] piece table (`Clx` -> `Pcdt` -> `PlcPcd`) parsing, text decoding,
 * and rebuilding for the legacy binary `.doc` editor.
 *
 * A `.doc`'s main document text is not one contiguous run: the `WordDocument`
 * stream holds the bytes, but the piece table maps document-wide character
 * positions ("CPs") onto byte ranges ("FCs") in that stream, each range
 * ("piece") independently either Windows-1252-compressed (1 byte/char) or
 * UTF-16LE (2 bytes/char). See `ole-document-doc-cp1252.ts` for the
 * byte<->code-point conversion this module calls into.
 *
 * @module ole-document-doc-pieces
 */
import { decodeCp1252, encodeCp1252 } from './ole-document-doc-cp1252.js';

/** One piece: a CP range mapped onto a byte range of a given encoding. */
export interface DocPiece {
	cpStart: number;
	cpEnd: number;
	/** Absolute byte offset into `WordDocument` of this piece's first character. */
	fc: number;
	compressed: boolean;
	/** PCD's own 2-byte flags word and 2-byte `prm`, kept verbatim for round-tripping an unmodified piece. */
	flagsWord: number;
	prm: number;
}

const CLXT_PRC = 1;
const CLXT_PCDT = 2;

/** Parse the `Clx` structure at `(fc, lcb)` in a table stream into its piece list, in CP order. */
export function parsePieceTable(
	tableStream: Uint8Array,
	clx: { fc: number; lcb: number },
): DocPiece[] {
	return parseDocClx(tableStream, clx).pieces;
}

/** CLX formatting records are opaque; retain their order because Prm1 refers to their indices. */
export interface DocClx {
	pieces: DocPiece[];
	/** Complete serialized Prc records preceding the Pcdt, including type/length headers. */
	prcBytes: Uint8Array;
}

/** Read a bounded CLX, retaining its formatting records ([MS-DOC] 2.9.38, 2.9.216). */
export function parseDocClx(tableStream: Uint8Array, clx: { fc: number; lcb: number }): DocClx {
	const view = new DataView(tableStream.buffer, tableStream.byteOffset, tableStream.byteLength);
	if (!Number.isInteger(clx.fc) || !Number.isInteger(clx.lcb) || clx.fc < 0 || clx.lcb < 5 ||
		clx.fc > tableStream.length - clx.lcb) throw new Error('Clx is outside the table stream');
	let off = clx.fc;
	const end = clx.fc + clx.lcb;
	let prcCount = 0;
	while (off < end) {
		const clxt = view.getUint8(off);
		if (clxt === CLXT_PRC) {
			if (off + 3 > end) throw new Error('Truncated Prc header');
			const cb = view.getUint16(off + 1, true);
			if (cb > end - off - 3) throw new Error('Prc extends beyond Clx');
			off += 1 + 2 + cb;
			prcCount++;
			continue;
		}
		if (clxt === CLXT_PCDT) {
			if (off + 5 > end) throw new Error('Truncated Pcdt header');
			const lcb = view.getInt32(off + 1, true);
			if (lcb < 4 || (lcb - 4) % 12 !== 0 || lcb !== end - off - 5)
				throw new Error('Invalid PlcPcd length');
			const pieces = parsePlcPcd(view, off + 5, lcb);
			for (const piece of pieces) {
				if ((piece.prm & 1) !== 0 && (piece.prm >>> 1) >= prcCount)
					throw new Error('Prm1 references a missing Prc');
			}
			return { pieces, prcBytes: tableStream.slice(clx.fc, off) };
		}
		throw new Error(`Unrecognised Clx entry type ${clxt} at ${off}`);
	}
	throw new Error('Clx has no Pcdt (piece table)');
}

function parsePlcPcd(view: DataView, plcStart: number, lcb: number): DocPiece[] {
	const n = Math.floor((lcb - 4) / (4 + 8));
	const cps: number[] = [];
	for (let i = 0; i <= n; i++) {
		const cp = view.getInt32(plcStart + i * 4, true);
		if ((i === 0 && cp !== 0) || (i > 0 && cp <= cps[i - 1]!))
			throw new Error('Piece CPs must start at zero and increase');
		cps.push(cp);
	}
	const pcdStart = plcStart + (n + 1) * 4;
	const pieces: DocPiece[] = [];
	for (let i = 0; i < n; i++) {
		const pOff = pcdStart + i * 8;
		const flagsWord = view.getUint16(pOff, true);
		const fcRaw = view.getUint32(pOff + 2, true);
		if ((fcRaw & 0x80000000) !== 0) throw new Error('Reserved FcCompressed bit is set');
		const prm = view.getUint16(pOff + 6, true);
		const compressed = ((fcRaw >>> 30) & 0x1) === 1;
		const fcField = fcRaw & 0x3fffffff;
		pieces.push({
			cpStart: cps[i]!,
			cpEnd: cps[i + 1]!,
			fc: compressed ? fcField >>> 1 : fcField,
			compressed,
			flagsWord,
			prm,
		});
	}
	return pieces;
}

/** Decode a piece's own character range to a string, reading its bytes from the `WordDocument` stream. */
function decodePieceText(wordDoc: Uint8Array, piece: DocPiece): string {
	const charLen = piece.cpEnd - piece.cpStart;
	const byteLen = charLen * (piece.compressed ? 1 : 2);
	if (!Number.isInteger(charLen) || charLen < 0 || !Number.isInteger(piece.fc) || piece.fc < 0 ||
		piece.fc > wordDoc.length - byteLen) throw new Error('Piece text is outside WordDocument');
	if (piece.compressed) {
		return decodeCp1252(wordDoc.subarray(piece.fc, piece.fc + charLen));
	}
	const view = new DataView(wordDoc.buffer, wordDoc.byteOffset, wordDoc.byteLength);
	let out = '';
	for (let i = 0; i < charLen; i++) {
		out += String.fromCharCode(view.getUint16(piece.fc + i * 2, true));
	}
	return out;
}

/** Decode text, optionally limited to the first maxChars CPs; validate all referenced text ranges. */
export function decodePiecesText(wordDoc: Uint8Array, pieces: readonly DocPiece[], maxChars?: number): string {
	if (maxChars !== undefined && (!Number.isInteger(maxChars) || maxChars < 0))
		throw new Error('Invalid text character limit');
	const out: string[] = [];
	for (const piece of pieces) {
		const bytes = (piece.cpEnd - piece.cpStart) * (piece.compressed ? 1 : 2);
		if (piece.cpEnd < piece.cpStart || piece.fc < 0 || piece.fc > wordDoc.length - bytes)
			throw new Error('Piece text is outside WordDocument');
		const end = Math.min(piece.cpEnd, maxChars ?? piece.cpEnd);
		if (end > piece.cpStart) out.push(decodePieceText(wordDoc, { ...piece, cpEnd: end }));
	}
	return out.join('');
}

/** Bytes for one new compressed or UTF-16LE piece's text, and which encoding was used. */
export function encodePieceText(text: string): { bytes: Uint8Array; compressed: boolean } {
	const compressedBytes = encodeCp1252(text);
	if (compressedBytes) {
		return { bytes: compressedBytes, compressed: true };
	}
	const bytes = new Uint8Array(text.length * 2);
	const view = new DataView(bytes.buffer);
	for (let i = 0; i < text.length; i++) {
		view.setUint16(i * 2, text.charCodeAt(i), true);
	}
	return { bytes, compressed: false };
}

/**
 * Replace the CP range `[cpStart, cpEnd)` (spanning one or more whole or
 * partial existing pieces) with one new piece, splitting any piece that only
 * partially overlaps the range and dropping the rest of the overlapped
 * content. Returns the new piece list with CPs renumbered from 0, in order.
 */
export function replacePieceRange(
	pieces: readonly DocPiece[],
	cpStart: number,
	cpEnd: number,
	newPiece: { fc: number; compressed: boolean; charLength: number; flagsWord: number; prm?: number },
): DocPiece[] {
	const result: Array<Omit<DocPiece, 'cpStart' | 'cpEnd'> & { charLength: number }> = [];
	let inserted = false;
	for (const piece of pieces) {
		const pieceCharLen = piece.cpEnd - piece.cpStart;
		const unit = piece.compressed ? 1 : 2;
		if (piece.cpEnd <= cpStart || piece.cpStart >= cpEnd) {
			result.push({ ...piece, charLength: pieceCharLen });
			continue;
		}
		// Overlaps the replaced range: keep the left remainder (if any), drop
		// the middle, insert the new piece once (at the first overlap), keep
		// the right remainder (if any).
		if (piece.cpStart < cpStart) {
			result.push({
				fc: piece.fc,
				compressed: piece.compressed,
				flagsWord: piece.flagsWord,
				prm: piece.prm,
				charLength: cpStart - piece.cpStart,
			});
		}
		if (!inserted) {
			result.push({
				fc: newPiece.fc,
				compressed: newPiece.compressed,
				flagsWord: newPiece.flagsWord,
				prm: newPiece.prm ?? 0,
				charLength: newPiece.charLength,
			});
			inserted = true;
		}
		if (piece.cpEnd > cpEnd) {
			const skipped = cpEnd - piece.cpStart;
			result.push({
				fc: piece.fc + skipped * unit,
				compressed: piece.compressed,
				flagsWord: piece.flagsWord,
				prm: piece.prm,
				charLength: piece.cpEnd - cpEnd,
			});
		}
	}
	if (!inserted) {
		result.push({
			fc: newPiece.fc,
			compressed: newPiece.compressed,
			flagsWord: newPiece.flagsWord,
			prm: newPiece.prm ?? 0,
			charLength: newPiece.charLength,
		});
	}

	let cp = 0;
	return result.map((r) => {
		const cpStartOut = cp;
		cp += r.charLength;
		return {
			cpStart: cpStartOut,
			cpEnd: cp,
			fc: r.fc,
			compressed: r.compressed,
			flagsWord: r.flagsWord,
			prm: r.prm,
		};
	});
}

/** Serialize pieces and optional verbatim Prc records back into a CLX. */
export function buildClxBytes(pieces: readonly DocPiece[], prcBytes: Uint8Array = new Uint8Array()): Uint8Array {
	const n = pieces.length;
	const plcLcb = (n + 1) * 4 + n * 8;
	const bytes = new Uint8Array(1 + 4 + plcLcb);
	const view = new DataView(bytes.buffer);
	view.setUint8(0, CLXT_PCDT);
	view.setInt32(1, plcLcb, true);
	let off = 5;
	for (const piece of pieces) {
		view.setInt32(off, piece.cpStart, true);
		off += 4;
	}
	view.setInt32(off, pieces[n - 1]?.cpEnd ?? 0, true);
	off += 4;
	for (const piece of pieces) {
		view.setUint16(off, piece.flagsWord, true);
		const fcField = piece.compressed ? piece.fc * 2 + 0x40000000 : piece.fc;
		view.setUint32(off + 2, fcField, true);
		view.setUint16(off + 6, piece.prm, true);
		off += 8;
	}
	const result = new Uint8Array(prcBytes.length + bytes.length);
	result.set(prcBytes);
	result.set(bytes, prcBytes.length);
	// Validate both caller-supplied records and references before writing them.
	parseDocClx(result, { fc: 0, lcb: result.length });
	return result;
}
