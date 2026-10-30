/** `readDocObjects`: pictures, shapes, fields, bookmarks and embedded objects of a
 * legacy Word document, positioned by absolute CP so a formatting or text reader
 * can be composed with the result. */
import { DocResourceLimitError } from './ole-document-doc-pieces.js';
import { resolveDocLimits } from './doc-limits.js';
import { openDocObjects } from './doc-objects-context.js';
import { readDocBookmarks } from './doc-bookmarks.js';
import { readDocFields } from './doc-fields.js';
import { readOleObjects } from './doc-ole-objects.js';
import { readInlinePictures } from './doc-pictures.js';
import { readDocDrawing, readFloatingShapes } from './doc-shapes.js';
import type { BlipBudget } from './doc-blip.js';
import type { DocObjects, DocObjectsLimits } from './doc-objects-types.js';

function guarded<T>(warnings: string[], label: string, fallback: T, run: () => T): T {
	try { return run(); }
	catch (error) {
		if (error instanceof DocResourceLimitError) throw error;
		warnings.push(`${label} could not be read: ${error instanceof Error ? error.message : 'invalid data'}`);
		return fallback;
	}
}

/**
 * Read the object model of a `.doc` file. Returns undefined when the bytes are not
 * a readable unencrypted Word 97-2003 document. Throws DocResourceLimitError when a
 * table exceeds `limits`; damage inside one optional structure is skipped and listed
 * in `warnings` instead.
 */
export function readDocObjects(input: Uint8Array, limits?: DocObjectsLimits): DocObjects | undefined {
	resolveDocLimits(limits); // invalid limits are a caller error, not a damaged file
	let ctx;
	try { ctx = openDocObjects(input, limits); }
	catch (error) {
		if (error instanceof DocResourceLimitError) throw error;
		return undefined;
	}
	if (!ctx) return undefined;
	const budget: BlipBudget = { maxBlipBytes: ctx.limits.maxBlipBytes, maxTotalBlipBytes: ctx.limits.maxTotalBlipBytes, used: 0 };
	const warnings = ctx.warnings;
	const fields = guarded(warnings, 'Fields', [], () => readDocFields(ctx));
	const ole = guarded(warnings, 'Embedded objects', { objects: [], previewCps: new Set<number>() }, () => readOleObjects(ctx, fields, budget));
	const inlinePictures = guarded(warnings, 'Inline pictures', [], () => readInlinePictures(ctx, budget, ole.previewCps));
	const drawing = guarded(warnings, 'Drawing group', { blips: [], shapes: new Map() }, () => readDocDrawing(ctx, budget));
	const floatingShapes = guarded(warnings, 'Floating shapes', [], () => readFloatingShapes(ctx, drawing));
	const bookmarks = guarded(warnings, 'Bookmarks', [], () => readDocBookmarks(ctx));
	return Object.freeze({
		stories: ctx.stories, inlinePictures, floatingShapes, fields, bookmarks, oleObjects: ole.objects,
		blips: drawing.blips, warnings,
	});
}
