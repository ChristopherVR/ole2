/** Inline pictures: 0x01 characters whose CHPX points (sprmCPicLocation) at a
 * PICFAndOfficeArtData in the Data stream ([MS-DOC] 2.9.192, 2.9.193). */
import { DocResourceLimitError } from './ole-document-doc-pieces.js';
import { SPRM, findSprm, sprmFlag, sprmUint } from './doc-objects-context.js';
import type { DocObjectsContext } from './doc-objects-context.js';
import { ART, artChildren, artPropertyText, artFixed, readArtProperties, readArtRecord } from './doc-officeart.js';
import type { ArtProperties, ArtRecord } from './doc-officeart.js';
import { decodeStoreEntry } from './doc-blip.js';
import type { BlipBudget } from './doc-blip.js';
import type { DocBlip, DocCrop, DocInlinePicture } from './doc-objects-types.js';

const PICF_SIZE = 0x44;

/** Crop fractions from OfficeArt cropFromTop/Bottom/Left/Right (0x100-0x103), when any is non-zero. */
export function cropFromProperties(props: ArtProperties): DocCrop | undefined {
	const [top, bottom, left, right] = [0x100, 0x101, 0x102, 0x103].map((id) => artFixed(props, id) ?? 0) as [number, number, number, number];
	return top || bottom || left || right ? { top, bottom, left, right } : undefined;
}

export type InlinePictureBody = Omit<DocInlinePicture, 'cp' | 'story'>;

/** Decode the PICFAndOfficeArtData at a Data-stream offset. */
export function readPicfAt(ctx: DocObjectsContext, offset: number, budget: BlipBudget): InlinePictureBody | undefined {
	const data = ctx.data;
	if (!data || offset < 0 || offset > data.length - PICF_SIZE) return undefined;
	const v = new DataView(data.buffer, data.byteOffset, data.byteLength);
	const lcb = v.getInt32(offset, true), cbHeader = v.getUint16(offset + 4, true), mm = v.getUint16(offset + 6, true);
	if (lcb < PICF_SIZE || lcb > data.length - offset || cbHeader !== PICF_SIZE) return undefined;
	const end = offset + lcb;
	const dxaGoal = v.getInt16(offset + 28, true), dyaGoal = v.getInt16(offset + 30, true);
	const mx = v.getUint16(offset + 32, true), my = v.getUint16(offset + 34, true);
	let at = offset + PICF_SIZE;
	let name: string | undefined;
	if (mm === 0x66) {
		if (at >= end) return undefined;
		const n = data[at]!;
		if (at + 1 + n > end) return undefined;
		name = n ? String.fromCharCode(...data.subarray(at + 1, at + 1 + n)) : undefined;
		at += 1 + n;
	}
	let props: ArtProperties | undefined;
	let blip: DocBlip | undefined;
	const blocks: ArtRecord[] = [];
	let bid: number | undefined;
	while (at + 8 <= end) {
		const rec = readArtRecord(data, at, end);
		if (!rec) break;
		if (rec.type === ART.spContainer && !props) {
			for (const child of artChildren(data, rec, ctx.limits.maxRecords)) {
				if (child.type === ART.fopt) {
					props = readArtProperties(data, child);
					for (const id of props.blipIds) if (id === 0x104) bid = props.values.get(id);
				}
			}
		} else blocks.push(rec);
		at = rec.end;
		if (blocks.length > ctx.limits.maxRecords) throw new DocResourceLimitError('Picture file-block count exceeds the processing limit');
	}
	// pib is a one-based index into the file blocks following the shape; a single block is used when absent.
	const pick = bid !== undefined && bid >= 1 && bid <= blocks.length ? bid : blocks.length ? 1 : 0;
	const rec = pick ? blocks[pick - 1] : undefined;
	if (rec) blip = decodeStoreEntry(data, rec, pick, ctx.word, budget);
	return {
		dataOffset: offset, dxaGoal, dyaGoal, mx, my,
		widthTwips: Math.round((dxaGoal * mx) / 1000), heightTwips: Math.round((dyaGoal * my) / 1000),
		crop: props ? cropFromProperties(props) : undefined, blip, name,
		alt: props ? artPropertyText(props, 0x381) : undefined,
	};
}

/** True when the character at `cp` is an object/picture placeholder with a plain picture location. */
export function pictureLocationAt(ctx: DocObjectsContext, cp: number): number | undefined {
	const sprms = ctx.sprmsAt(cp);
	const loc = findSprm(sprms, SPRM.picLocation);
	if (!loc || !sprmFlag(sprms, SPRM.fSpec) || sprmFlag(sprms, SPRM.fData) || sprmFlag(sprms, SPRM.fOle2)) return undefined;
	return sprmUint(loc);
}

/** Every inline picture in all stories, excluding the given preview CPs (they belong to embedded objects). */
export function readInlinePictures(ctx: DocObjectsContext, budget: BlipBudget, exclude: ReadonlySet<number> = new Set()): DocInlinePicture[] {
	const out: DocInlinePicture[] = [];
	const cache = new Map<number, InlinePictureBody | undefined>();
	for (let cp = ctx.text.indexOf('\u0001'); cp >= 0; cp = ctx.text.indexOf('\u0001', cp + 1)) {
		if (exclude.has(cp)) continue;
		const loc = pictureLocationAt(ctx, cp);
		if (loc === undefined) continue;
		if (out.length >= ctx.limits.maxRecords) throw new DocResourceLimitError('Inline picture count exceeds the processing limit');
		if (!cache.has(loc)) {
			try { cache.set(loc, readPicfAt(ctx, loc, budget)); }
			catch (error) { if (error instanceof DocResourceLimitError) throw error; cache.set(loc, undefined); }
		}
		const body = cache.get(loc);
		if (body) out.push({ cp, story: ctx.storyOf(cp), ...body });
		else ctx.warnings.push(`Inline picture at cp ${cp} is unreadable`);
	}
	return out;
}
