/** Bounds-checked OfficeArt ([MS-ODRAW] 2.2) record walking, property tables and
 * shape-container collection. Every length is validated against its parent. */
import { DocResourceLimitError } from './ole-document-doc-pieces.js';

export interface ArtRecord {
	readonly ver: number;
	readonly inst: number;
	readonly type: number;
	/** Offset of the 8-byte header. */
	readonly start: number;
	/** Offset of the body and the end of the record (exclusive). */
	readonly body: number;
	readonly end: number;
}

export const ART = {
	dggContainer: 0xf000, bstore: 0xf001, dgContainer: 0xf002, spgrContainer: 0xf003, spContainer: 0xf004,
	fdggBlock: 0xf006, fbse: 0xf007, fsp: 0xf00a, fopt: 0xf00b, clientTextbox: 0xf00d, clientAnchor: 0xf010,
	clientData: 0xf011, tertiaryFopt: 0xf122, secondaryFopt: 0xf121,
} as const;

/** Read the record header at `off`; undefined when it does not fit in `[off, limit)`. */
export function readArtRecord(bytes: Uint8Array, off: number, limit: number): ArtRecord | undefined {
	if (off < 0 || limit > bytes.length || off + 8 > limit) return undefined;
	const w = bytes[off]! | (bytes[off + 1]! << 8);
	const type = bytes[off + 2]! | (bytes[off + 3]! << 8);
	const len = (bytes[off + 4]! | (bytes[off + 5]! << 8) | (bytes[off + 6]! << 16) | (bytes[off + 7]! << 24)) >>> 0;
	if (len > limit - off - 8) return undefined;
	return { ver: w & 0xf, inst: w >>> 4, type, start: off, body: off + 8, end: off + 8 + len };
}

/** Child records of a container, stopping at the first malformed header. */
export function artChildren(bytes: Uint8Array, parent: ArtRecord, max: number): ArtRecord[] {
	const out: ArtRecord[] = [];
	let at = parent.body;
	while (at + 8 <= parent.end) {
		const rec = readArtRecord(bytes, at, parent.end);
		if (!rec) break;
		if (out.length >= max) throw new DocResourceLimitError('OfficeArt record count exceeds the processing limit');
		out.push(rec);
		at = rec.end;
	}
	return out;
}

export interface ArtProperties {
	/** Property id (flag bits cleared) to its 4-byte operand. */
	readonly values: Map<number, number>;
	/** Property ids whose operand is a BLIP id (fBid). */
	readonly blipIds: Set<number>;
	/** Complex property payloads, by id. */
	readonly complex: Map<number, Uint8Array>;
}

/** Parse an OfficeArtFOPT body ([MS-ODRAW] 2.2.9). */
export function readArtProperties(bytes: Uint8Array, rec: ArtRecord): ArtProperties {
	const values = new Map<number, number>(), blipIds = new Set<number>(), complex = new Map<number, Uint8Array>();
	const count = rec.inst;
	const tableEnd = rec.body + count * 6;
	if (tableEnd > rec.end) return { values, blipIds, complex };
	let extra = tableEnd;
	for (let i = 0; i < count; i++) {
		const at = rec.body + i * 6;
		const opid = bytes[at]! | (bytes[at + 1]! << 8);
		const op = ((bytes[at + 2]! | (bytes[at + 3]! << 8) | (bytes[at + 4]! << 16) | (bytes[at + 5]! << 24)) >>> 0);
		const id = opid & 0x3fff;
		if (opid & 0x8000) {
			const len = op;
			if (len <= rec.end - extra) complex.set(id, bytes.subarray(extra, extra + len));
			extra += Math.min(len, rec.end - extra);
			values.set(id, len);
		} else {
			values.set(id, op);
			if (opid & 0x4000) blipIds.add(id);
		}
	}
	return { values, blipIds, complex };
}

export interface ArtShape {
	readonly spid: number;
	readonly shapeType: number;
	readonly flags: number;
	readonly properties: ArtProperties;
	readonly hasTextbox: boolean;
	readonly isGroup: boolean;
}

function mergeProps(a: ArtProperties, b: ArtProperties): ArtProperties {
	for (const [k, v] of b.values) a.values.set(k, v);
	for (const k of b.blipIds) a.blipIds.add(k);
	for (const [k, v] of b.complex) a.complex.set(k, v);
	return a;
}

/** Read one OfficeArtSpContainer. */
export function readArtShape(bytes: Uint8Array, sp: ArtRecord, maxRecords: number): ArtShape | undefined {
	let spid = 0, shapeType = 0, flags = 0, found = false, hasTextbox = false;
	let properties: ArtProperties = { values: new Map(), blipIds: new Set(), complex: new Map() };
	for (const child of artChildren(bytes, sp, maxRecords)) {
		if (child.type === ART.fsp && child.end - child.body >= 8) {
			spid = (bytes[child.body]! | (bytes[child.body + 1]! << 8) | (bytes[child.body + 2]! << 16) | (bytes[child.body + 3]! << 24)) >>> 0;
			flags = (bytes[child.body + 4]! | (bytes[child.body + 5]! << 8) | (bytes[child.body + 6]! << 16) | (bytes[child.body + 7]! << 24)) >>> 0;
			shapeType = child.inst;
			found = true;
		} else if (child.type === ART.fopt || child.type === ART.secondaryFopt || child.type === ART.tertiaryFopt) {
			properties = mergeProps(properties, readArtProperties(bytes, child));
		} else if (child.type === ART.clientTextbox) hasTextbox = true;
	}
	return found ? { spid, shapeType, flags, properties, hasTextbox, isGroup: (flags & 1) !== 0 } : undefined;
}

/** Collect every shape under a drawing (or group) container, depth- and count-limited. */
export function collectArtShapes(
	bytes: Uint8Array, container: ArtRecord, limits: { maxRecords: number; maxArtDepth: number }, out: ArtShape[] = [], depth = 0,
): ArtShape[] {
	if (depth > limits.maxArtDepth) throw new DocResourceLimitError('OfficeArt nesting exceeds the processing limit');
	for (const child of artChildren(bytes, container, limits.maxRecords)) {
		if (child.type === ART.spContainer) {
			const shape = readArtShape(bytes, child, limits.maxRecords);
			if (shape) {
				if (out.length >= limits.maxRecords) throw new DocResourceLimitError('OfficeArt shape count exceeds the processing limit');
				out.push(shape);
			}
		} else if (child.type === ART.spgrContainer || child.type === ART.dgContainer) {
			collectArtShapes(bytes, child, limits, out, depth + 1);
		}
	}
	return out;
}

/** UTF-16LE complex property text without the terminating NUL. */
export function artPropertyText(props: ArtProperties, id: number): string | undefined {
	const raw = props.complex.get(id);
	if (!raw || raw.length < 2) return undefined;
	let s = '';
	for (let i = 0; i + 1 < raw.length; i += 2) {
		const c = raw[i]! | (raw[i + 1]! << 8);
		if (c === 0) break;
		s += String.fromCharCode(c);
	}
	return s || undefined;
}

/** 16.16 fixed-point property as a number. */
export function artFixed(props: ArtProperties, id: number): number | undefined {
	const v = props.values.get(id);
	return v === undefined ? undefined : (v | 0) / 65536;
}

/** msocolor property to 0xRRGGBB; undefined for scheme or system colours that need a palette. */
export function artColor(props: ArtProperties, id: number): number | undefined {
	const v = props.values.get(id);
	if (v === undefined) return undefined;
	if ((v >>> 24) & 0x18) return undefined;
	return ((v & 0xff) << 16) | (v & 0xff00) | ((v >>> 16) & 0xff);
}

/** Boolean-group property (use bits in the high half). */
export function artBool(props: ArtProperties, id: number, bit: number): boolean | undefined {
	const v = props.values.get(id);
	if (v === undefined || ((v >>> (bit + 16)) & 1) === 0) return undefined;
	return ((v >>> bit) & 1) === 1;
}
