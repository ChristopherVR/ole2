/** Explicit scalar creation using the verified existing-row insertion path. */
import { createXlsNumericCellWorkbookStream, type XlsCreateNumericCellFailure } from './legacy-excel-create-numeric-cell.js';
import { editXlsStringWorkbookStream, type XlsPreservedStringFailure } from './legacy-excel-preserved-string-edit.js';
import { editXlsBoolErrorWorkbookStream, normalizeXlsErrorValue } from './legacy-excel-bool-error-edit.js';
import type { XlsCellValue } from './legacy-excel-workbook.js';
import { OLE_MAGIC } from './ole2-parser-types.js';
import { readCompoundFileStream } from './ole2-stream-edit.js';
import { resizeCompoundFileStream } from './ole2-stream-resize.js';

export type XlsCreatableCellValue = Exclude<XlsCellValue, null>;
export interface XlsCreateCell { row: number; col: number; value: XlsCreatableCellValue; xf: number; worksheetIndex?: number }
export type XlsCreateCellFailure = XlsCreateNumericCellFailure | XlsPreservedStringFailure;
export type XlsCreateCellResult = { status: 'edited'; bytes: Uint8Array; recalculationRequired: true }
	| { status: 'unchanged'; bytes: Uint8Array; reason: XlsCreateCellFailure };

/** Bare BIFF8 Workbook stream. Staged insertion and conversion are never
 * exposed separately: any refusal returns the original input. */
export function createXlsCellWorkbookStream(input: Uint8Array, supplied: XlsCreateCell): XlsCreateCellResult {
	const unchanged = (reason: XlsCreateCellFailure): XlsCreateCellResult => ({ status: 'unchanged', bytes: input, reason });
	const { row, col, xf, worksheetIndex } = supplied;
	const candidate = supplied.value;
	const value = typeof candidate === 'object' ? normalizeXlsErrorValue(candidate) : candidate;
	if (value === undefined || (typeof value !== 'number' && typeof value !== 'string' && typeof value !== 'boolean' && typeof value !== 'object')) return unchanged('invalid-edit');
	if (typeof value === 'string' && (value.length > 32767 || /[\uD800-\uDBFF](?![\uDC00-\uDFFF])|(?:^|[^\uD800-\uDBFF])[\uDC00-\uDFFF]/u.test(value))) return unchanged('invalid-edit');
	const edit = { row, col, xf, worksheetIndex };
	const insertion = createXlsNumericCellWorkbookStream(input, { ...edit, value: typeof value === 'number' ? value : 0 });
	if (insertion.status !== 'edited' || typeof value === 'number') return insertion;
	const converted = typeof value === 'string'
		? editXlsStringWorkbookStream(insertion.bytes, { ...edit, value })
		: editXlsBoolErrorWorkbookStream(insertion.bytes, { ...edit, value });
	return converted.status === 'edited' ? converted : unchanged(converted.reason);
}

/** Whole CFB or bare BIFF8 input. Wraps the completed edit only once. */
export function createXlsCell(input: Uint8Array, edit: XlsCreateCell): XlsCreateCellResult {
	if (!OLE_MAGIC.every((value, index) => input[index] === value)) return createXlsCellWorkbookStream(input, edit);
	try {
		const streams = ['Workbook', 'Book'].map(name => ({ name, bytes: readCompoundFileStream(input, [name]) })).filter(stream => stream.bytes !== undefined);
		if (streams.length !== 1) return { status: 'unchanged', bytes: input, reason: 'unsupported-workbook' };
		const stream = streams[0]!, result = createXlsCellWorkbookStream(stream.bytes!, edit);
		if (result.status !== 'edited') return { ...result, bytes: input };
		const resized = resizeCompoundFileStream(input, [stream.name], result.bytes);
		return resized.ok ? { ...result, bytes: resized.bytes } : { status: 'unchanged', bytes: input, reason: 'container-not-writable' };
	} catch { return { status: 'unchanged', bytes: input, reason: 'malformed-records' }; }
}
