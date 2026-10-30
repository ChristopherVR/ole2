/** Bounded zlib/DEFLATE decoder (RFC 1950/1951) so metafile BLIPs decompress
 * without a platform dependency. Output growth is capped before allocation. */

export class DocInflateError extends Error {
	constructor(message: string) {
		super(message);
		this.name = 'DocInflateError';
	}
}

const LEN_BASE = [3, 4, 5, 6, 7, 8, 9, 10, 11, 13, 15, 17, 19, 23, 27, 31, 35, 43, 51, 59, 67, 83, 99, 115, 131, 163, 195, 227, 258];
const LEN_EXTRA = [0, 0, 0, 0, 0, 0, 0, 0, 1, 1, 1, 1, 2, 2, 2, 2, 3, 3, 3, 3, 4, 4, 4, 4, 5, 5, 5, 5, 0];
const DIST_BASE = [1, 2, 3, 4, 5, 7, 9, 13, 17, 25, 33, 49, 65, 97, 129, 193, 257, 385, 513, 769, 1025, 1537, 2049, 3073, 4097, 6145, 8193, 12289, 16385, 24577];
const DIST_EXTRA = [0, 0, 0, 0, 1, 1, 2, 2, 3, 3, 4, 4, 5, 5, 6, 6, 7, 7, 8, 8, 9, 9, 10, 10, 11, 11, 12, 12, 13, 13];
const CL_ORDER = [16, 17, 18, 0, 8, 7, 9, 6, 10, 5, 11, 4, 12, 3, 13, 2, 14, 1, 15];

interface Huffman { counts: Uint16Array; symbols: Uint16Array }

function build(lengths: ArrayLike<number>, n: number): Huffman {
	const counts = new Uint16Array(16), symbols = new Uint16Array(n);
	for (let i = 0; i < n; i++) counts[lengths[i]!]!++;
	const offs = new Uint16Array(16);
	for (let i = 1; i < 15; i++) offs[i + 1] = offs[i]! + counts[i]!;
	for (let i = 0; i < n; i++) if (lengths[i]) symbols[offs[lengths[i]!]!++] = i;
	return { counts, symbols };
}

let fixedLit: Huffman | undefined, fixedDist: Huffman | undefined;

/** Inflate raw DEFLATE data starting at `start`; throws DocInflateError past `maxOutput` or on corrupt input. */
export function inflateRaw(src: Uint8Array, start: number, maxOutput: number): Uint8Array {
	let pos = start, bitBuf = 0, bitCnt = 0;
	let out = new Uint8Array(Math.min(Math.max(1024, (src.length - start) * 3), maxOutput + 1));
	let outLen = 0;
	const bits = (n: number): number => {
		while (bitCnt < n) {
			if (pos >= src.length) throw new DocInflateError('Truncated DEFLATE stream');
			bitBuf |= src[pos++]! << bitCnt;
			bitCnt += 8;
		}
		const v = bitBuf & ((1 << n) - 1);
		bitBuf >>>= n;
		bitCnt -= n;
		return v;
	};
	const decode = (h: Huffman): number => {
		let code = 0, first = 0, index = 0;
		for (let len = 1; len <= 15; len++) {
			code |= bits(1);
			const count = h.counts[len]!;
			if (code - count < first) return h.symbols[index + (code - first)]!;
			index += count; first += count; first <<= 1; code <<= 1;
		}
		throw new DocInflateError('Invalid Huffman code');
	};
	const put = (b: number) => {
		if (outLen >= maxOutput) throw new DocInflateError('Inflated size exceeds the processing limit');
		if (outLen >= out.length) { const next = new Uint8Array(Math.min(out.length * 2, maxOutput + 1)); next.set(out); out = next; }
		out[outLen++] = b;
	};
	let last = 0;
	while (!last) {
		last = bits(1);
		const type = bits(2);
		if (type === 0) {
			bitBuf = 0; bitCnt = 0;
			if (pos + 4 > src.length) throw new DocInflateError('Truncated stored block');
			const len = src[pos]! | (src[pos + 1]! << 8);
			if ((len ^ 0xffff) !== (src[pos + 2]! | (src[pos + 3]! << 8))) throw new DocInflateError('Bad stored block length');
			pos += 4;
			if (pos + len > src.length) throw new DocInflateError('Truncated stored block');
			for (let i = 0; i < len; i++) put(src[pos + i]!);
			pos += len;
			continue;
		}
		if (type === 3) throw new DocInflateError('Invalid DEFLATE block type');
		let lit: Huffman, dist: Huffman;
		if (type === 1) {
			if (!fixedLit) {
				const l = new Uint8Array(288);
				l.fill(8, 0, 144); l.fill(9, 144, 256); l.fill(7, 256, 280); l.fill(8, 280, 288);
				fixedLit = build(l, 288);
				fixedDist = build(new Uint8Array(30).fill(5), 30);
			}
			lit = fixedLit; dist = fixedDist!;
		} else {
			const nlen = bits(5) + 257, ndist = bits(5) + 1, ncode = bits(4) + 4;
			if (nlen > 286 || ndist > 30) throw new DocInflateError('Bad dynamic block counts');
			const cl = new Uint8Array(19);
			for (let i = 0; i < ncode; i++) cl[CL_ORDER[i]!] = bits(3);
			const clh = build(cl, 19);
			const lengths = new Uint8Array(nlen + ndist);
			for (let i = 0; i < nlen + ndist;) {
				const sym = decode(clh);
				if (sym < 16) { lengths[i++] = sym; continue; }
				let prev = 0, rep: number;
				if (sym === 16) { if (i === 0) throw new DocInflateError('Bad length repeat'); prev = lengths[i - 1]!; rep = 3 + bits(2); }
				else if (sym === 17) rep = 3 + bits(3);
				else rep = 11 + bits(7);
				if (i + rep > nlen + ndist) throw new DocInflateError('Bad length repeat');
				while (rep--) lengths[i++] = prev;
			}
			lit = build(lengths.subarray(0, nlen), nlen);
			dist = build(lengths.subarray(nlen), ndist);
		}
		for (;;) {
			const sym = decode(lit);
			if (sym < 256) { put(sym); continue; }
			if (sym === 256) break;
			const li = sym - 257;
			if (li >= 29) throw new DocInflateError('Bad length symbol');
			const len = LEN_BASE[li]! + bits(LEN_EXTRA[li]!);
			const ds = decode(dist);
			if (ds >= 30) throw new DocInflateError('Bad distance symbol');
			const d = DIST_BASE[ds]! + bits(DIST_EXTRA[ds]!);
			if (d > outLen) throw new DocInflateError('Distance before start of output');
			for (let i = 0; i < len; i++) put(out[outLen - d]!);
		}
	}
	return out.slice(0, outLen);
}

/** Inflate a zlib stream (2-byte header, DEFLATE data; the Adler-32 trailer is not enforced). */
export function inflateZlib(src: Uint8Array, maxOutput: number): Uint8Array {
	if (src.length < 2 || (src[0]! & 0x0f) !== 8 || ((src[0]! << 8) | src[1]!) % 31 !== 0 || (src[1]! & 0x20))
		throw new DocInflateError('Not a zlib stream');
	return inflateRaw(src, 2, maxOutput);
}
