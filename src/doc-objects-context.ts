/** Opened DOC streams shared by the object readers: FIB, piece table, all-story
 * text, story ranges and tolerant character-property lookup by CP. */
import { unwrapDocBytes } from './ole-document-doc-cfb.js';
import { readDocFib, readFcLcbAt } from './ole-document-doc-fib.js';
import type { FcLcb, ParsedDocFib } from './ole-document-doc-fib.js';
import { extractChpxBlob, parseBteTable } from './ole-document-doc-fkp.js';
import type { BteTable } from './ole-document-doc-fkp.js';
import { DocResourceLimitError, decodePiecesText, parseDocClx } from './ole-document-doc-pieces.js';
import type { DocPiece } from './ole-document-doc-pieces.js';
import { readCompoundFileStream } from './ole2-stream-edit.js';
import { DEFAULT_DOC_OBJECTS_LIMITS } from './doc-objects-types.js';
import type { DocObjectsLimits, DocStoryKind, DocStoryRange } from './doc-objects-types.js';

export interface DocSprmRecord {
	readonly opcode: number;
	readonly operand: Uint8Array;
}

/** Character SPRM opcodes used by the object readers ([MS-DOC] 2.6.1). */
export const SPRM = { fSpec: 0x0855, fData: 0x0806, fOle2: 0x080a, picLocation: 0x6a03, objLocation: 0x680e } as const;

/** FibRgFcLcb97 indices used by the object readers ([MS-DOC] 2.5.6). */
export const FCLCB = {
	fldMom: 16, fldHdr: 17, fldFtn: 18, fldAtn: 19, sttbfBkmk: 21, plcfBkf: 22, plcfBkl: 23,
	spaMom: 40, spaHdr: 41, fldEdn: 48, dggInfo: 50, fldTxbx: 57, fldHdrTxbx: 59,
} as const;

export interface DocObjectsContext {
	readonly input: Uint8Array;
	readonly word: Uint8Array;
	readonly table: Uint8Array;
	readonly data: Uint8Array | undefined;
	readonly fib: ParsedDocFib;
	readonly pieces: readonly DocPiece[];
	/** Text of every story, concatenated, indexed by absolute CP. */
	readonly text: string;
	readonly stories: readonly DocStoryRange[];
	readonly limits: Readonly<Required<DocObjectsLimits>>;
	readonly warnings: string[];
	fcLcb(index: number): FcLcb;
	storyOf(cp: number): DocStoryKind;
	/** Direct character SPRMs at a CP: CHPX exceptions followed by piece PRM exceptions. */
	sprmsAt(cp: number): readonly DocSprmRecord[];
}

const SPRA_SIZE = [1, 1, 2, 4, 2, 2, 0, 3] as const;

/** Walk a grpprl, stopping quietly at the first unparseable record. */
export function parseSprms(bytes: Uint8Array, start: number, end: number, max = 1024): DocSprmRecord[] {
	const out: DocSprmRecord[] = [];
	let at = start;
	while (at + 2 <= end && out.length < max) {
		const opcode = bytes[at]! | (bytes[at + 1]! << 8);
		const spra = opcode >>> 13;
		let size: number = SPRA_SIZE[spra]!;
		let head = 2;
		if (spra === 6) {
			if (at + 3 > end) break;
			if (opcode === 0xd608) { size = (bytes[at + 2]! | (bytes[at + 3]! << 8)) - 1; head = 4; }
			else size = bytes[at + 2]!, head = 3;
			if (size < 0) break;
		}
		if (at + head + size > end) break;
		out.push({ opcode, operand: bytes.subarray(at + head, at + head + size) });
		at += head + size;
	}
	return out;
}

export function sprmUint(r: DocSprmRecord): number {
	let v = 0;
	for (let i = Math.min(r.operand.length, 4) - 1; i >= 0; i--) v = v * 256 + r.operand[i]!;
	return v;
}

export function findSprm(sprms: readonly DocSprmRecord[], opcode: number): DocSprmRecord | undefined {
	for (let i = sprms.length - 1; i >= 0; i--) if (sprms[i]!.opcode === opcode) return sprms[i];
	return undefined;
}

/** Toggle operand: 1 is on; 0x80/0x81 refer to the style and count as absent. */
export function sprmFlag(sprms: readonly DocSprmRecord[], opcode: number): boolean {
	const r = findSprm(sprms, opcode);
	return r !== undefined && r.operand[0] === 1;
}

const STORY_ORDER: readonly [DocStoryKind, number][] = [
	['main', 3], ['footnotes', 4], ['headers', 5], ['comments', 7], ['endnotes', 8], ['textboxes', 9], ['headerTextboxes', 10],
];

export function resolveLimits(limits?: DocObjectsLimits): Required<DocObjectsLimits> {
	const out = { ...DEFAULT_DOC_OBJECTS_LIMITS, ...Object.fromEntries(Object.entries(limits ?? {}).filter(([, v]) => v !== undefined)) } as Required<DocObjectsLimits>;
	for (const v of Object.values(out)) if (!Number.isSafeInteger(v) || v < 0) throw new Error('Invalid DOC object limits');
	return out;
}

/** Open a DOC for object reading; undefined when it is not a readable unencrypted Word 97-2003 file. */
export function openDocObjects(input: Uint8Array, limitsIn?: DocObjectsLimits): DocObjectsContext | undefined {
	const limits = resolveLimits(limitsIn);
	const cfb = unwrapDocBytes(input);
	if (!cfb) return undefined;
	const { wordDocBytes: word, tableBytes: table } = cfb;
	const fib = readDocFib(word);
	const view = new DataView(word.buffer, word.byteOffset, word.byteLength);
	const lwBase = fib.fibRgFcLcbOffset - 2 - 22 * 4;
	const ccp = (i: number) => view.getInt32(lwBase + i * 4, true);
	const pieces = parseDocClx(table, fib.clx, limits.maxPieces).pieces;
	const total = pieces.at(-1)?.cpEnd ?? 0;
	if (total > limits.maxCharacters) throw new DocResourceLimitError('DOC text exceeds the processing limit');
	const text = decodePiecesText(word, pieces, total);
	const stories: DocStoryRange[] = [];
	let base = 0;
	for (const [story, idx] of STORY_ORDER) {
		const len = ccp(idx);
		if (len < 0) throw new Error('Negative story length');
		const end = Math.min(base + len, total);
		if (len > 0) stories.push(Object.freeze({ story, cpStart: base, cpEnd: end }));
		base += len;
	}
	const warnings: string[] = [];
	let chpx: BteTable | undefined;
	try {
		if (fib.plcfbteChpx.lcb > 65_536 * 8 + 4) throw new DocResourceLimitError('CHPX table exceeds the processing limit');
		chpx = fib.plcfbteChpx.lcb >= 4 ? parseBteTable(table, fib.plcfbteChpx) : undefined;
	} catch (error) { warnings.push(`Character formatting table unreadable: ${(error as Error).message}`); }
	const prcs = prcGrpprls(table, fib.clx);
	const sprmCache = new Map<number, readonly DocSprmRecord[]>();
	const data = readCompoundFileStream(input, ['Data']);
	const ctx: DocObjectsContext = {
		input, word, table, data, fib, pieces, text, stories, limits, warnings,
		fcLcb: (index) => readFcLcbAt(word, fib, index),
		storyOf(cp) {
			for (const s of stories) if (cp >= s.cpStart && cp < s.cpEnd) return s.story;
			return 'main';
		},
		sprmsAt(cp) {
			const cached = sprmCache.get(cp);
			if (cached) return cached;
			const out: DocSprmRecord[] = [];
			const piece = findPiece(pieces, cp);
			if (piece) {
				const fc = piece.fc + (cp - piece.cpStart) * (piece.compressed ? 1 : 2);
				if (chpx) {
					try {
						const blob = extractChpxBlob(word, chpx, fc);
						if (blob.length > 1) out.push(...parseSprms(blob, 1, blob.length));
					} catch { /* uncovered or damaged page: no direct properties */ }
				}
				if ((piece.prm & 1) === 1) {
					const g = prcs[piece.prm >>> 1];
					if (g) out.push(...parseSprms(g, 0, g.length));
				}
			}
			if (sprmCache.size < 65_536) sprmCache.set(cp, out);
			return out;
		},
	};
	return ctx;
}

function findPiece(pieces: readonly DocPiece[], cp: number): DocPiece | undefined {
	let lo = 0, hi = pieces.length - 1;
	while (lo <= hi) {
		const mid = (lo + hi) >>> 1, p = pieces[mid]!;
		if (cp < p.cpStart) hi = mid - 1; else if (cp >= p.cpEnd) lo = mid + 1; else return p;
	}
	return undefined;
}

/** Grpprl of each Prc in the CLX (used by Prm1). Damaged entries end the list. */
function prcGrpprls(table: Uint8Array, clx: FcLcb): Uint8Array[] {
	const out: Uint8Array[] = [];
	let off = clx.fc;
	const end = Math.min(clx.fc + clx.lcb, table.length);
	while (off + 3 <= end && table[off] === 1 && out.length < 65_536) {
		const cb = table[off + 1]! | (table[off + 2]! << 8);
		if (off + 3 + cb > end) break;
		out.push(table.subarray(off + 3, off + 3 + cb));
		off += 3 + cb;
	}
	return out;
}

/** Bounds-checked sub-range of the table stream for a PLC-like structure. */
export function tableRange(ctx: DocObjectsContext, at: FcLcb): Uint8Array | undefined {
	if (at.lcb === 0) return undefined;
	if (at.fc > ctx.table.length || at.lcb > ctx.table.length - at.fc) {
		ctx.warnings.push('A table-stream structure lies outside the table stream');
		return undefined;
	}
	return ctx.table.subarray(at.fc, at.fc + at.lcb);
}
