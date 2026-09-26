import { readCompoundFileStream, replaceCompoundFileStream } from './ole2-stream-edit.js';
import {
	readSummaryProperties,
	writeSummaryTextProperty,
	SUMMARY_STREAM_NAME,
} from './ole-summary-properties.js';
import type { SummaryProperties, SummaryTextProperty } from './ole-summary-properties.js';

function summaryStream(bytes: Uint8Array): Uint8Array | undefined {
	return readCompoundFileStream(bytes, [SUMMARY_STREAM_NAME]);
}

/** Read standard document text metadata, when present. Does not interpret application content. */
export function readLegacyOfficeMetadata(bytes: Uint8Array): SummaryProperties | undefined {
	const stream = summaryStream(bytes);
	return stream && readSummaryProperties(stream);
}

/** Update one existing root-level text property in its allocated space, preserving the rest of the compound file.
 * Applies to legacy DOC/XLS/PPT/VSD/PUB containers carrying SummaryInformation.
 * Returns the exact original input for unsupported edits; does not add missing properties or expand streams.
 */
export function writeLegacyOfficeMetadata(
	bytes: Uint8Array,
	key: SummaryTextProperty,
	value: string,
): Uint8Array {
	const stream = summaryStream(bytes);
	if (!stream) return bytes;
	const replacement = writeSummaryTextProperty(stream, key, value);
	return replacement === stream
		? bytes
		: replaceCompoundFileStream(bytes, [SUMMARY_STREAM_NAME], replacement);
}
