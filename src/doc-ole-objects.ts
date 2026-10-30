/** Embedded OLE objects: characters carrying sprmCFOle2 whose sprmCPicLocation
 * is the ObjectPool storage id ("_" + decimal), with the ProgID read from the
 * storage's CompObj stream and the preview picture from the following 0x01
 * character ([MS-DOC] 2.9.173 sprmCFOle2, 2.9.158 sprmCObjLocation, MS-OLEDS 2.3.7). */
import { parseOle2 } from './ole2-parser-read.js';
import type { Ole2File } from './ole2-parser-types.js';
import { DocResourceLimitError } from './ole-document-doc-pieces.js';
import { SPRM, findSprm, sprmFlag, sprmUint } from './doc-objects-context.js';
import type { DocObjectsContext } from './doc-objects-context.js';
import { pictureLocationAt, readPicfAt } from './doc-pictures.js';
import type { BlipBudget } from './doc-blip.js';
import type { DocField, DocInlinePicture, DocOleObject } from './doc-objects-types.js';

function guid(b: Uint8Array): string | undefined {
	if (b.length < 16 || b.every((x) => x === 0)) return undefined;
	const h = (s: number, e: number, rev: boolean) => {
		const part = Array.from(b.subarray(s, e)).map((x) => x.toString(16).padStart(2, '0'));
		return (rev ? part.reverse() : part).join('');
	};
	return `${h(0, 4, true)}-${h(4, 6, true)}-${h(6, 8, true)}-${h(8, 10, false)}-${h(10, 16, false)}`;
}

/** ProgID from a CompObj stream: header, AnsiUserType, AnsiClipboardFormat, then the ProgID string. */
export function readCompObjProgId(s: Uint8Array): string | undefined {
	if (s.length < 32) return undefined;
	const v = new DataView(s.buffer, s.byteOffset, s.byteLength);
	let at = 28;
	const lengthString = (): string | undefined | null => {
		if (at + 4 > s.length) return null;
		const len = v.getUint32(at, true);
		at += 4;
		if (len > 4096 || at + len > s.length) return null;
		const text = String.fromCharCode(...s.subarray(at, at + (len > 0 && s[at + len - 1] === 0 ? len - 1 : len)));
		at += len;
		return text || undefined;
	};
	if (lengthString() === null) return undefined;
	if (at + 4 > s.length) return undefined;
	const marker = v.getUint32(at, true);
	if (marker === 0xffffffff || marker === 0xfffffffe) at += 8;
	else if (marker === 0) at += 4;
	else if (lengthString() === null) return undefined;
	const progId = lengthString();
	return progId ?? undefined;
}

export function readOleObjects(
	ctx: DocObjectsContext, fields: readonly DocField[], budget: BlipBudget,
): { objects: DocOleObject[]; previewCps: Set<number> } {
	const objects: DocOleObject[] = [];
	const previewCps = new Set<number>();
	let cfb: Ole2File | undefined;
	const text = ctx.text;
	for (const ch of ['\u0014', '\u0001', '\u0013']) {
		for (let cp = text.indexOf(ch); cp >= 0; cp = text.indexOf(ch, cp + 1)) {
			const sprms = ctx.sprmsAt(cp);
			if (!sprmFlag(sprms, SPRM.fOle2)) continue;
			const loc = findSprm(sprms, SPRM.picLocation) ?? findSprm(sprms, SPRM.objLocation);
			if (!loc) continue;
			if (objects.length >= ctx.limits.maxRecords) throw new DocResourceLimitError('Embedded object count exceeds the processing limit');
			const storageName = `_${sprmUint(loc)}`;
			if (!cfb) {
				try { cfb = parseOle2(ctx.input); } catch { ctx.warnings.push('Compound file directory is unreadable'); return { objects, previewCps }; }
			}
			const storage = cfb.entries.find((e) => e.type === 1 && e.path?.length === 2 && e.path[0]!.toLowerCase() === 'objectpool' && e.path[1] === storageName);
			if (!storage) { ctx.warnings.push(`Object storage ${storageName} at cp ${cp} is missing`); continue; }
			const streams = cfb.entries.filter((e) => e.parentId === storage.id && e.type === 2).map((e) => e.name);
			let progId: string | undefined;
			try {
				const compObj = cfb.getStreamByPath(['ObjectPool', storageName, '\u0001CompObj']);
				progId = compObj ? readCompObjProgId(compObj) : undefined;
			} catch { /* optional stream */ }
			let preview: DocInlinePicture | undefined;
			const next = text.charCodeAt(cp + 1) === 1 ? cp + 1 : -1;
			const previewLoc = next >= 0 ? pictureLocationAt(ctx, next) : undefined;
			if (previewLoc !== undefined) {
				previewCps.add(next);
				try {
					const body = readPicfAt(ctx, previewLoc, budget);
					if (body) preview = { cp: next, story: ctx.storyOf(next), ...body };
				} catch (error) { if (error instanceof DocResourceLimitError) throw error; }
			}
			const enclosing = fields.filter((f) => f.cpBegin < cp && f.cpEnd > cp).sort((a, b) => b.depth - a.depth)[0];
			objects.push({
				cp, story: ctx.storyOf(cp), storageName, progId, clsid: guid(storage.clsid), streams, preview,
				linked: /^\s*LINK\b/iu.test(enclosing?.instruction ?? ''),
			});
		}
	}
	return { objects: objects.sort((a, b) => a.cp - b.cp), previewCps };
}
