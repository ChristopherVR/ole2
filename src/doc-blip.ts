/** OfficeArt BLIP decoding ([MS-ODRAW] 2.2.23-2.2.32): PNG, JPEG, DIB, TIFF and
 * compressed or plain metafiles, plus the FBSE store entries that wrap them. */
import { ART, artChildren, readArtRecord } from './doc-officeart.js';
import type { ArtRecord } from './doc-officeart.js';
import { inflateZlib } from './doc-inflate.js';
import type { DocBlip, DocBlipFormat } from './doc-objects-types.js';

const BLIP_FORMATS: Readonly<Record<number, { format: DocBlipFormat; mime: string }>> = {
	0xf01a: { format: 'emf', mime: 'image/emf' },
	0xf01b: { format: 'wmf', mime: 'image/wmf' },
	0xf01c: { format: 'pict', mime: 'image/x-pict' },
	0xf01d: { format: 'jpeg', mime: 'image/jpeg' },
	0xf02a: { format: 'jpeg', mime: 'image/jpeg' },
	0xf01e: { format: 'png', mime: 'image/png' },
	0xf01f: { format: 'bmp', mime: 'image/bmp' },
	0xf029: { format: 'tiff', mime: 'image/tiff' },
};

export interface BlipBudget { maxBlipBytes: number; maxTotalBlipBytes: number; used: number }

export function isBlipRecord(rec: ArtRecord): boolean {
	return BLIP_FORMATS[rec.type] !== undefined;
}

function hex(bytes: Uint8Array): string {
	let s = '';
	for (const b of bytes) s += b.toString(16).padStart(2, '0');
	return s;
}

function u32(b: Uint8Array, o: number): number {
	return (b[o]! | (b[o + 1]! << 8) | (b[o + 2]! << 16) | (b[o + 3]! << 24)) >>> 0;
}

/** Pixel size of PNG, JPEG, BMP and TIFF-less headers; undefined when unreadable. */
export function imagePixelSize(data: Uint8Array, format: DocBlipFormat): [number, number] | undefined {
	if (format === 'png' && data.length >= 24 && data[1] === 0x50) {
		return [((data[16]! << 24) | (data[17]! << 16) | (data[18]! << 8) | data[19]!) >>> 0, ((data[20]! << 24) | (data[21]! << 16) | (data[22]! << 8) | data[23]!) >>> 0];
	}
	if (format === 'bmp' && data.length >= 26) return [u32(data, 18), Math.abs(u32(data, 22) | 0)];
	if (format === 'jpeg') {
		let at = 2;
		while (at + 9 < data.length) {
			if (data[at] !== 0xff) { at++; continue; }
			const marker = data[at + 1]!;
			if (marker === 0xff) { at++; continue; }
			if (marker >= 0xc0 && marker <= 0xcf && marker !== 0xc4 && marker !== 0xc8 && marker !== 0xcc)
				return [(data[at + 7]! << 8) | data[at + 8]!, (data[at + 5]! << 8) | data[at + 6]!];
			if (marker === 0xd8 || marker === 0x01 || (marker >= 0xd0 && marker <= 0xd7)) { at += 2; continue; }
			at += 2 + ((data[at + 2]! << 8) | data[at + 3]!);
		}
	}
	return undefined;
}

/** Wrap a headerless DIB (BITMAPINFOHEADER and pixels) in a BITMAPFILEHEADER. */
function dibToBmp(dib: Uint8Array): Uint8Array | undefined {
	if (dib.length < 40) return undefined;
	const headerSize = u32(dib, 0);
	if (headerSize < 40 || headerSize > dib.length) return undefined;
	const bpp = dib[14]! | (dib[15]! << 8), compression = u32(dib, 16);
	let colors = u32(dib, 32);
	if (bpp <= 8 && colors === 0) colors = 1 << bpp;
	if (colors > 65_536) return undefined;
	const palette = colors * 4;
	const masks = headerSize === 40 && (compression === 3 || compression === 6) ? (compression === 3 ? 12 : 16) : 0;
	const off = 14 + headerSize + masks + palette;
	const out = new Uint8Array(14 + dib.length);
	out[0] = 0x42; out[1] = 0x4d;
	new DataView(out.buffer).setUint32(2, out.length, true);
	new DataView(out.buffer).setUint32(10, Math.min(off, out.length), true);
	out.set(dib, 14);
	return out;
}

/** Decode one BLIP record; undefined for unsupported, truncated or over-budget images. */
export function decodeBlip(bytes: Uint8Array, rec: ArtRecord, index: number, budget: BlipBudget, delayed: boolean): DocBlip | undefined {
	const kind = BLIP_FORMATS[rec.type];
	if (!kind) return undefined;
	let at = rec.body;
	const uidLen = (rec.inst & 1) === 1 ? 32 : 16;
	if (at + uidLen > rec.end) return undefined;
	const uid = hex(bytes.subarray(at, at + 16));
	at += uidLen;
	let data: Uint8Array;
	let metafile: DocBlip['metafile'];
	if (rec.type <= 0xf01c) {
		if (at + 34 > rec.end) return undefined;
		const cbSize = u32(bytes, at), cbSave = u32(bytes, at + 28);
		const compression = bytes[at + 32]!;
		const rc = [0, 4, 8, 12].map((o) => u32(bytes, at + 4 + o) | 0);
		metafile = { boundsEmu: [rc[0]!, rc[1]!, rc[2]!, rc[3]!], sizeEmu: [u32(bytes, at + 20) | 0, u32(bytes, at + 24) | 0] };
		at += 34;
		if (cbSave > rec.end - at || cbSize > budget.maxBlipBytes) return undefined;
		if (compression !== 0 && compression !== 0xfe) return undefined;
		const stored = bytes.subarray(at, at + cbSave);
		try {
			data = compression === 0 ? inflateZlib(stored, Math.min(budget.maxBlipBytes, Math.max(cbSize, 1))) : stored;
		} catch { return undefined; }
	} else {
		at += 1; // tag
		if (at > rec.end) return undefined;
		data = bytes.subarray(at, rec.end);
		if (rec.type === 0xf01f) {
			const bmp = data.length <= budget.maxBlipBytes ? dibToBmp(data) : undefined;
			if (!bmp) return undefined;
			data = bmp;
		}
	}
	if (data.length > budget.maxBlipBytes || budget.used + data.length > budget.maxTotalBlipBytes) return undefined;
	budget.used += data.length;
	const size = imagePixelSize(data, kind.format);
	return { index, format: kind.format, mime: kind.mime, data: data.slice(), widthPx: size?.[0], heightPx: size?.[1], uid, ...(metafile ? { metafile } : {}), delayed };
}

/** Decode an FBSE (or a bare BLIP) store entry; delayed BLIPs are read from the WordDocument stream. */
export function decodeStoreEntry(bytes: Uint8Array, rec: ArtRecord, index: number, word: Uint8Array | undefined, budget: BlipBudget): DocBlip | undefined {
	if (isBlipRecord(rec)) return decodeBlip(bytes, rec, index, budget, false);
	if (rec.type !== ART.fbse || rec.end - rec.body < 36) return undefined;
	const cbName = bytes[rec.body + 33]!;
	const sizeField = u32(bytes, rec.body + 20), foDelay = u32(bytes, rec.body + 28);
	const embedded = rec.body + 36 + cbName;
	if (embedded + 8 <= rec.end) {
		const inner = readArtRecord(bytes, embedded, rec.end);
		if (inner && isBlipRecord(inner)) return decodeBlip(bytes, inner, index, budget, false);
	}
	if (sizeField > 0 && foDelay !== 0xffffffff && word && foDelay < word.length) {
		const inner = readArtRecord(word, foDelay, word.length);
		if (inner && isBlipRecord(inner)) return decodeBlip(word, inner, index, budget, true);
	}
	return undefined;
}

/** Decode every entry of an OfficeArtBStoreContainer, keeping index alignment (undefined holes are skipped, indices stay one-based). */
export function decodeBStore(bytes: Uint8Array, bstore: ArtRecord, word: Uint8Array | undefined, budget: BlipBudget, max: number): DocBlip[] {
	const out: DocBlip[] = [];
	artChildren(bytes, bstore, max).forEach((rec, i) => {
		const blip = decodeStoreEntry(bytes, rec, i + 1, word, budget);
		if (blip) out.push(blip);
	});
	return out;
}
