import { parseOle2 } from './ole2-parser-read.js';
import { ENTRY_TYPE_STREAM } from './ole2-parser-types.js';

const VISIO_DOCUMENT_STREAM = 'visiodocument';
const TRAILER_STREAM_TYPE = 0x14;
const HEADER_POINTER_OFFSET = 0x24;
const VISIO_HEADER = 'Visio (TM) Drawing\r\n';

/**
 * Structural facts read from a legacy Visio binary drawing. This is not a
 * drawing parser and does not claim that shapes or pages were decoded.
 * Header layout follows Apache POI and LibreOffice's independent readers:
 * https://github.com/apache/poi/blob/trunk/poi-scratchpad/src/main/java/org/apache/poi/hdgf/HDGFDiagram.java
 * https://github.com/apache/poi/blob/trunk/poi-scratchpad/src/main/java/org/apache/poi/hdgf/pointers/PointerFactory.java
 * https://github.com/LibreOffice/libvisio/blob/master/src/lib/VSDParser.cpp
 */
export interface LegacyVisioInspection {
	format: 'vsd';
	version: number;
	documentSize: number;
	streams: string[];
	trailerType: number;
	trailerOffset: number;
	trailerLength: number;
	trailerFormat: number;
	trailerCompressed: boolean;
	reason: string;
}

/** Validate VSD's signature, version-dependent pointer and trailer range. */
export function inspectLegacyVisio(
	input: Uint8Array | ArrayBuffer,
): LegacyVisioInspection | undefined {
	const bytes = input instanceof Uint8Array ? input : new Uint8Array(input);
	try {
		const buffer = bytes.buffer.slice(
			bytes.byteOffset,
			bytes.byteOffset + bytes.byteLength,
		) as ArrayBuffer;
		const ole = parseOle2(buffer);
		const streamEntries = ole.entries.filter((entry) => entry.type === ENTRY_TYPE_STREAM);
		const names = streamEntries.map((entry) => entry.name);
		const streamName = names.find((name) => name.toLowerCase() === VISIO_DOCUMENT_STREAM);
		if (!streamName) return undefined;
		const stream = ole.getStream(streamName);
		if (!stream || stream.length < HEADER_POINTER_OFFSET + 16) return undefined;

		const view = new DataView(stream.buffer, stream.byteOffset, stream.byteLength);
		if (new TextDecoder('windows-1252').decode(stream.subarray(0, 20)) !== VISIO_HEADER)
			return undefined;
		const version = view.getUint16(0x1a, true);
		const documentSize = view.getUint32(0x1c, true);
		if (version < 5 || documentSize < HEADER_POINTER_OFFSET + 16 || documentSize > stream.length)
			return undefined;

		// V5 uses a 16-byte pointer; V6+ uses an 18-byte pointer.
		const v5 = version === 5;
		const trailerType = v5
			? view.getUint16(HEADER_POINTER_OFFSET, true)
			: view.getUint32(HEADER_POINTER_OFFSET, true);
		const trailerFormat = view.getUint16(HEADER_POINTER_OFFSET + (v5 ? 2 : 16), true);
		const trailerOffset = view.getUint32(HEADER_POINTER_OFFSET + 8, true);
		const trailerLength = view.getUint32(HEADER_POINTER_OFFSET + 12, true);
		if (trailerType !== TRAILER_STREAM_TYPE || trailerLength === 0) return undefined;
		if (trailerOffset < HEADER_POINTER_OFFSET + (v5 ? 16 : 18)) return undefined;
		if (trailerOffset > documentSize || trailerLength > documentSize - trailerOffset)
			return undefined;

		return {
			format: 'vsd',
			version,
			documentSize,
			streams: names,
			trailerType,
			trailerOffset,
			trailerLength,
			trailerFormat,
			trailerCompressed: (trailerFormat & 0x2) !== 0,
			reason: `Visio format version ${version} header and bounded TrailerStream pointer validated; drawing content was not parsed.`,
		};
	} catch {
		return undefined;
	}
}
