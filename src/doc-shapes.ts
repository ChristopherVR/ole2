/** Floating shapes and pictures: PlcSpaMom/PlcSpaHdr (FSPA) anchors joined with
 * the OfficeArt drawing group in the Table stream (OfficeArtContent, DggInfo). */
import { DocResourceLimitError } from './ole-document-doc-pieces.js';
import { FCLCB, tableRange } from './doc-objects-context.js';
import type { DocObjectsContext } from './doc-objects-context.js';
import { ART, artBool, artChildren, artColor, artFixed, artPropertyText, collectArtShapes, readArtRecord } from './doc-officeart.js';
import type { ArtShape } from './doc-officeart.js';
import { decodeBStore } from './doc-blip.js';
import type { BlipBudget } from './doc-blip.js';
import { cropFromProperties } from './doc-pictures.js';
import type { DocBlip, DocFloatingShape, DocShapeAnchorRelative, DocShapeWrap } from './doc-objects-types.js';

const FSPA_SIZE = 26;
const WRAP: readonly DocShapeWrap[] = ['unknown', 'topAndBottom', 'square', 'none', 'tight', 'through'];
const REL_X: readonly DocShapeAnchorRelative[] = ['column', 'margin', 'page'];
const REL_Y: readonly DocShapeAnchorRelative[] = ['margin', 'page', 'paragraph'];
const WRAP_SIDE = ['both', 'left', 'right', 'largest'] as const;
const MSOSPT_PICTURE_FRAME = 75;

export interface DocDrawing {
	readonly blips: DocBlip[];
	/** Shapes by spid, across every drawing (main and header). */
	readonly shapes: Map<number, ArtShape>;
}

/** Parse OfficeArtContent: the drawing group (BLIP store) and every drawing's shapes. */
export function readDocDrawing(ctx: DocObjectsContext, budget: BlipBudget): DocDrawing {
	const empty: DocDrawing = { blips: [], shapes: new Map() };
	const bytes = tableRange(ctx, ctx.pair(FCLCB.dggInfo));
	if (!bytes) return empty;
	const dgg = readArtRecord(bytes, 0, bytes.length);
	if (!dgg || dgg.type !== ART.dggContainer) {
		ctx.warnings.push('Drawing group is not an OfficeArt container');
		return empty;
	}
	const blips: DocBlip[] = [];
	for (const child of artChildren(bytes, dgg, ctx.limits.maxRecords)) {
		if (child.type === ART.bstore) blips.push(...decodeBStore(bytes, child, ctx.word, budget, ctx.limits.maxRecords));
	}
	const shapes = new Map<number, ArtShape>();
	let at = dgg.end;
	while (at + 9 <= bytes.length) {
		const dg = readArtRecord(bytes, at + 1, bytes.length); // one-byte dgglbl precedes each drawing
		if (!dg || dg.type !== ART.dgContainer) break;
		for (const shape of collectArtShapes(bytes, dg, ctx.limits)) {
			if ((shape.flags & 4) === 0 && !shapes.has(shape.spid)) shapes.set(shape.spid, shape); // skip patriarch
		}
		at = dg.end;
	}
	return { blips, shapes };
}

function shapeFrom(
	ctx: DocObjectsContext, drawing: DocDrawing, v: DataView, at: number, cp: number, story: DocFloatingShape['story'],
): DocFloatingShape {
	const spid = v.getUint32(at, true);
	const left = v.getInt32(at + 4, true), top = v.getInt32(at + 8, true), right = v.getInt32(at + 12, true), bottom = v.getInt32(at + 16, true);
	const flags = v.getUint16(at + 20, true);
	const art = drawing.shapes.get(spid);
	const props = art?.properties;
	const pib = props && props.blipIds.has(0x104) ? props.values.get(0x104) : undefined;
	const blip = pib !== undefined ? drawing.blips.find((b) => b.index === pib) : undefined;
	const filled = props ? artBool(props, 0x1bf, 4) : undefined;
	const lined = props ? artBool(props, 0x1ff, 3) : undefined;
	const rotation = props ? artFixed(props, 0x4) : undefined;
	const txid = props?.values.get(0x80);
	const opacity = props ? artFixed(props, 0x182) : undefined;
	const lineWidth = props?.values.get(0x1cb);
	return {
		spid, cp, story,
		leftTwips: left, topTwips: top, rightTwips: right, bottomTwips: bottom,
		widthTwips: right - left, heightTwips: bottom - top,
		relativeX: REL_X[(flags >> 1) & 3] ?? 'unknown', relativeY: REL_Y[(flags >> 3) & 3] ?? 'unknown',
		wrap: WRAP[(flags >> 5) & 15] ?? 'unknown', wrapSide: WRAP_SIDE[(flags >> 9) & 15] ?? 'unknown',
		behindText: (flags & 0x4000) !== 0, inHeader: (flags & 1) !== 0,
		textboxCount: v.getUint32(at + 22, true),
		shapeType: art?.shapeType, isPicture: art?.shapeType === MSOSPT_PICTURE_FRAME || pib !== undefined,
		blip, blipIndex: pib, crop: props ? cropFromProperties(props) : undefined,
		textboxId: txid !== undefined && txid !== 0 ? txid : undefined,
		fill: props ? { enabled: filled ?? true, colorRgb: artColor(props, 0x181), opacity } : undefined,
		line: props ? { enabled: lined ?? true, colorRgb: artColor(props, 0x1c0), widthEmu: lineWidth } : undefined,
		rotationDegrees: rotation, name: props ? artPropertyText(props, 0x380) : undefined,
		alt: props ? artPropertyText(props, 0x381) : undefined,
		properties: props ? new Map(props.values) : new Map(),
	};
}

/** Every FSPA anchor in the main and header stories, joined with its OfficeArt shape. */
export function readFloatingShapes(ctx: DocObjectsContext, drawing: DocDrawing): DocFloatingShape[] {
	const out: DocFloatingShape[] = [];
	for (const [index, storyBase] of [[FCLCB.spaMom, 0], [FCLCB.spaHdr, ctx.stories.find((s) => s.story === 'headers')?.cpStart ?? 0]] as const) {
		const bytes = tableRange(ctx, ctx.pair(index));
		if (!bytes) continue;
		if (bytes.length < 4 || (bytes.length - 4) % (4 + FSPA_SIZE) !== 0) {
			ctx.warnings.push('FSPA table has an invalid length');
			continue;
		}
		const n = (bytes.length - 4) / (4 + FSPA_SIZE);
		if (n + out.length > ctx.limits.maxRecords) throw new DocResourceLimitError('Shape count exceeds the processing limit');
		const v = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
		for (let i = 0; i < n; i++) {
			const cp = storyBase + v.getInt32(i * 4, true);
			if (cp < 0 || cp >= ctx.text.length) { ctx.warnings.push(`FSPA anchor ${i} lies outside the document`); continue; }
			out.push(shapeFrom(ctx, drawing, v, (n + 1) * 4 + i * FSPA_SIZE, cp, ctx.storyOf(cp)));
		}
	}
	return out;
}
