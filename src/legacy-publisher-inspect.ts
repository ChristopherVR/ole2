import { parseOle2 } from './ole2-parser-read.js';
import { ENTRY_TYPE_STREAM } from './ole2-parser-types.js';
import { readCompoundFileStream } from './ole2-stream-edit.js';

const CONTENTS_PATH = ['Contents'] as const;
const ESCHER_PATH = ['Escher', 'EscherStm'] as const;
const QUILL_PATH = ['Quill', 'QuillSub', 'CONTENTS'] as const;

/** The version is inferred only from the four-byte Contents signature. */
export type LegacyPublisherVersion = 'publisher-97-2000' | 'publisher-2002';

/**
 * Validated Publisher container facts. This does not decode page layout or
 * expose a rendering/editing API.
 *
 * Header signatures and the 2002 stream requirements are based on the
 * LibreOffice libmspub format gate:
 * https://github.com/LibreOffice/libmspub/blob/master/src/lib/MSPUBDocument.cpp
 */
export interface LegacyPublisherInspection {
	format: 'pub';
	version: LegacyPublisherVersion;
	streams: string[];
	validatedStreams: string[];
	reason: string;
}

/** Validate Publisher's Contents signature and version-specific stream set. */
export function inspectLegacyPublisher(
	input: Uint8Array | ArrayBuffer,
): LegacyPublisherInspection | undefined {
	const bytes = input instanceof Uint8Array ? input : new Uint8Array(input);
	try {
		const ole = parseOle2(
			bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength) as ArrayBuffer,
		);
		const streamEntries = ole.entries.filter((entry) => entry.type === ENTRY_TYPE_STREAM);
		const streams = streamEntries.map((entry) => entry.name);
		const contents = readCompoundFileStream(bytes, CONTENTS_PATH);
		if (!contents || contents.length < 4) return undefined;
		if (contents[0] !== 0xe8 || contents[1] !== 0xac || contents[3] !== 0) return undefined;

		const version =
			contents[2] === 0x22
				? 'publisher-97-2000'
				: contents[2] === 0x2c
					? 'publisher-2002'
					: undefined;
		if (!version) return undefined;

		const validatedStreams = ['Contents'];
		if (version === 'publisher-2002') {
			if (!readCompoundFileStream(bytes, ESCHER_PATH) || !readCompoundFileStream(bytes, QUILL_PATH))
				return undefined;
			validatedStreams.push('Escher/EscherStm', 'Quill/QuillSub/CONTENTS');
		}

		return {
			format: 'pub',
			version,
			streams,
			validatedStreams,
			reason: `Publisher ${version === 'publisher-2002' ? '2002' : '97/2000'} Contents signature validated; page content was not parsed.`,
		};
	} catch {
		return undefined;
	}
}
