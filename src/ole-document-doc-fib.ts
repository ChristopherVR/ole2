/* Adapted from ChristopherVR/pptx-viewer packages/core/src/core/utils, Apache-2.0. Original source: https://github.com/ChristopherVR/pptx-viewer. */
/**
 * [MS-DOC] File Information Block (FIB) field access for the `WordDocument`
 * stream of a legacy binary `.doc` (Word 97-2003, `nFib` 193/"Word 97") OLE
 * payload.
 *
 * Only the handful of fields `ole-document-doc-editor.ts` needs are exposed:
 * which table stream holds the piece table (`fWhichTblStm`), the main body
 * character count (`ccpText`), the `WordDocument` stream's meaningful byte
 * length (`cbMac`, distinct from the stream's padded physical size), and the
 * `(fc, lcb)` pointers into the table stream for the piece table (CLX), the
 * paragraph/character formatting bookkeeping tables (`PlcBtePapx`/
 * `PlcBteChpx`), and the section table (`PlcfSed`). Every offset is derived
 * from the FIB's own `csw`/`cslw`/`cbRgFcLcb` counts rather than hardcoded,
 * per [MS-DOC] 2.5.1-2.5.4, so this keeps working across `nFib` variants that
 * add trailing `FibRgFcLcb` entries.
 *
 * @module ole-document-doc-fib
 */

/** Index into `FibRgFcLcb97` for each `(fc, lcb)` pair this module reads/patches. See [MS-DOC] 2.5.3. */
const FC_LCB_INDEX = {
	plcfbteChpx: 12,
	plcfbtePapx: 13,
	sed: 6,
	clx: 33,
} as const;

/** One `(fc, lcb)` pointer into the table stream. */
export interface FcLcb {
	fc: number;
	lcb: number;
}

/** Byte offsets and current values of every FIB field this module touches. */
export interface DocFib {
	/** Effective format version, including nFibNew when present. */
	nFib: number;
	nFibBase: number;
	/** Byte offset of the `flags1` word (FibBase), for patching `fComplex`. */
	flags1Offset: number;
	flags1: number;
	/** Which table stream the piece table lives in. */
	tableStreamName: '0Table' | '1Table';
	/** Byte offset of `FibRgLw97.cbMac` (meaningful length of `WordDocument`). */
	cbMacOffset: number;
	cbMac: number;
	/** Byte offset of `FibRgLw97.ccpText` (main document character count). */
	ccpTextOffset: number;
	ccpText: number;
	/** Populated non-main stories require CP tables this editor cannot safely update. */
	ccpOtherStories: number;
	plcfbteChpx: FcLcb;
	plcfbtePapx: FcLcb;
	sed: FcLcb;
	clx: FcLcb;
	/** Byte offset of the `FibRgFcLcb97` array, for locating each pair's own offset when patching. */
	fibRgFcLcbOffset: number;
	fibRgFcLcbCount: number;
}

/** `fComplex` bit (bit 2) of `FibBase.flags1`. Set once this module has performed an incremental (piece-append) edit. */
const F_COMPLEX_BIT = 1 << 2;

/**
 * Read every FIB field {@link DocFib} exposes from a `WordDocument` stream's
 * bytes. Throws if the stream does not start with the expected `wIdent`
 * magic (0xA5EC): callers should already have gated on `WordDocument` stream
 * presence via `ole-payload-kind.ts` before calling this.
 */
export function readDocFib(wordDoc: Uint8Array): DocFib {
	const view = new DataView(wordDoc.buffer, wordDoc.byteOffset, wordDoc.byteLength);
	const wIdent = view.getUint16(0x0, true);
	if (wIdent !== 0xa5ec) {
		throw new Error('Not a WordDocument FIB (bad wIdent)');
	}
	const nFibBase = view.getUint16(2, true);
	if (![0x00c1, 0x00d9, 0x0101, 0x010c, 0x0112].includes(nFibBase))
		throw new Error('Unsupported Word FIB version');

	const flags1Offset = 0xa;
	const flags1 = view.getUint16(flags1Offset, true);
	if ((flags1 & 0x8100) !== 0) throw new Error('Encrypted or obfuscated Word document');
	const fWhichTblStm = (flags1 >> 9) & 0x1;

	const csw = view.getUint16(0x20, true);
	const fibRgW97Offset = 0x22;
	const cslwOffset = fibRgW97Offset + csw * 2;
	const cslw = view.getUint16(cslwOffset, true);
	const fibRgLw97Offset = cslwOffset + 2;
	const cbRgFcLcbCountOffset = fibRgLw97Offset + cslw * 4;
	const fibRgFcLcbOffset = cbRgFcLcbCountOffset + 2;
	const fibRgFcLcbCount = view.getUint16(cbRgFcLcbCountOffset, true);
	if (csw !== 14 || cslw !== 22 ||
		fibRgFcLcbOffset + fibRgFcLcbCount * 8 + 2 > wordDoc.length)
		throw new Error('Truncated Word FIB arrays');
	const cswNewOffset = fibRgFcLcbOffset + fibRgFcLcbCount * 8;
	const cswNew = view.getUint16(cswNewOffset, true);
	if (cswNewOffset + 2 + cswNew * 2 > wordDoc.length) throw new Error('Truncated Word FIB extension');
	const nFib = cswNew === 0 ? nFibBase : view.getUint16(cswNewOffset + 2, true);
	const expectedCounts = new Map([[0x00c1, 93], [0x00d9, 108], [0x0101, 136], [0x010c, 164], [0x0112, 183]]);
	if (expectedCounts.get(nFib) !== fibRgFcLcbCount ||
		(cswNew !== 0 && cswNew !== (nFib === 0x0112 ? 5 : 2)))
		throw new Error('Word FIB version and array counts disagree');
	const ccpText = view.getInt32(fibRgLw97Offset + 12, true);
	const otherStories = [4, 5, 7, 8, 9, 10].map((i) => view.getInt32(fibRgLw97Offset + i * 4, true));
	if (ccpText < 0 || otherStories.some((cp) => cp < 0)) throw new Error('Negative Word story length');

	function fcLcbPair(index: number): FcLcb {
		const off = fibRgFcLcbOffset + index * 8;
		return { fc: view.getUint32(off, true), lcb: view.getUint32(off + 4, true) };
	}

	const cbMacOffset = fibRgLw97Offset + 0;
	const ccpTextOffset = fibRgLw97Offset + 12;

	return {
		nFib,
		nFibBase,
		flags1Offset,
		flags1,
		tableStreamName: fWhichTblStm ? '1Table' : '0Table',
		cbMacOffset,
		cbMac: view.getUint32(cbMacOffset, true),
		ccpTextOffset,
		ccpText,
		ccpOtherStories: otherStories.reduce((sum, cp) => sum + cp, 0),
		plcfbteChpx: fcLcbPair(FC_LCB_INDEX.plcfbteChpx),
		plcfbtePapx: fcLcbPair(FC_LCB_INDEX.plcfbtePapx),
		sed: fcLcbPair(FC_LCB_INDEX.sed),
		clx: fcLcbPair(FC_LCB_INDEX.clx),
		fibRgFcLcbOffset,
		fibRgFcLcbCount,
	};
}

/** Absolute byte offset of one `FibRgFcLcb97` pair's `fc` field, for {@link patchDocFib}. */
function fcLcbFieldOffset(fib: DocFib, key: 'plcfbteChpx' | 'plcfbtePapx' | 'clx'): number {
	return fib.fibRgFcLcbOffset + FC_LCB_INDEX[key] * 8;
}

/**
 * Read an arbitrary `FibRgFcLcb97` pair by its [MS-DOC] 2.5.3 index, for
 * feature-detection checks that are not one of {@link DocFib}'s named fields
 * (`ole-document-doc-editor.ts`'s "does this doc use footnotes/annotations/
 * fields/bookmarks" gate).
 */
export function readFcLcbAt(wordDoc: Uint8Array, fib: DocFib, index: number): FcLcb {
	if (!Number.isInteger(index) || index < 0 || index >= fib.fibRgFcLcbCount)
		throw new Error('FIB fc/lcb index outside declared array');
	const view = new DataView(wordDoc.buffer, wordDoc.byteOffset, wordDoc.byteLength);
	const off = fib.fibRgFcLcbOffset + index * 8;
	return { fc: view.getUint32(off, true), lcb: view.getUint32(off + 4, true) };
}

/**
 * Apply in-place FIB field updates to a (mutable) copy of the `WordDocument`
 * stream's bytes: new `ccpText`, new `cbMac`, `fComplex` set, and the new
 * `(fc, lcb)` for `clx`/`plcfbteChpx`/`plcfbtePapx` (all three are
 * re-appended rather than resized in place; `PlcfSed` is patched separately
 * in-place by the caller since its size never changes for a same-section-count
 * edit).
 */
export function patchDocFib(
	wordDoc: Uint8Array,
	fib: DocFib,
	patch: {
		ccpText: number;
		cbMac: number;
		clx: FcLcb;
		plcfbteChpx: FcLcb;
		plcfbtePapx: FcLcb;
	},
): void {
	const view = new DataView(wordDoc.buffer, wordDoc.byteOffset, wordDoc.byteLength);
	view.setUint16(fib.flags1Offset, fib.flags1 | F_COMPLEX_BIT, true);
	view.setUint32(fib.cbMacOffset, patch.cbMac, true);
	view.setInt32(fib.ccpTextOffset, patch.ccpText, true);

	const clxOff = fcLcbFieldOffset(fib, 'clx');
	view.setUint32(clxOff, patch.clx.fc, true);
	view.setUint32(clxOff + 4, patch.clx.lcb, true);

	const chpxOff = fcLcbFieldOffset(fib, 'plcfbteChpx');
	view.setUint32(chpxOff, patch.plcfbteChpx.fc, true);
	view.setUint32(chpxOff + 4, patch.plcfbteChpx.lcb, true);

	const papxOff = fcLcbFieldOffset(fib, 'plcfbtePapx');
	view.setUint32(papxOff, patch.plcfbtePapx.fc, true);
	view.setUint32(papxOff + 4, patch.plcfbtePapx.lcb, true);
}
