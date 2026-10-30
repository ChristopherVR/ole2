/** Opened DOC streams for the object readers. The context itself is the shared
 * {@link StoryContext}; this module keeps the FIB indices and character sprm
 * opcodes the object readers use and the bounds-checked table-stream slice. */
import type { FcLcb } from './ole-document-doc-fib.js';
import { DocResourceLimitError } from './ole-document-doc-pieces.js';
import { openStoryContext } from './doc-story.js';
import type { StoryContext } from './doc-story.js';
import type { DocObjectsLimits } from './doc-objects-types.js';

export { findSprm, sprmFlag, sprmUint } from './doc-sprm.js';

/** The shared reading context; `fcLcb` is the same lookup as `pair`. */
export type DocObjectsContext = StoryContext;

/** Character SPRM opcodes used by the object readers ([MS-DOC] 2.6.1). */
export const SPRM = { fSpec: 0x0855, fData: 0x0806, fOle2: 0x080a, picLocation: 0x6a03, objLocation: 0x680e } as const;

/** FibRgFcLcb97 indices used by the object readers ([MS-DOC] 2.5.6). */
export const FCLCB = {
	fldMom: 16, fldHdr: 17, fldFtn: 18, fldAtn: 19, sttbfBkmk: 21, plcfBkf: 22, plcfBkl: 23,
	spaMom: 40, spaHdr: 41, fldEdn: 48, dggInfo: 50, fldTxbx: 57, fldHdrTxbx: 59,
} as const;

/** Open a DOC for object reading; undefined when it is not a readable unencrypted Word 97-2003 file. */
export function openDocObjects(input: Uint8Array, limits?: DocObjectsLimits): DocObjectsContext | undefined {
	try {
		return openStoryContext(input, limits);
	} catch (error) {
		if (error instanceof DocResourceLimitError || (error instanceof Error && /limits must be/.test(error.message))) throw error;
		return undefined;
	}
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
