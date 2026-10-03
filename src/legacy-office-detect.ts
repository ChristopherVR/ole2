import { parseOle2 } from './ole2-parser-read.js';
import { ENTRY_TYPE_STREAM } from './ole2-parser-types.js';
import { inspectLegacyVisio } from './legacy-visio-inspect.js';
import { inspectLegacyPublisher } from './legacy-publisher-inspect.js';

export type LegacyOfficeFormat = 'doc' | 'xls' | 'ppt' | 'visio' | 'publisher' | 'unknown';
export type LegacyOfficeSupport = 'partial' | 'unsupported' | 'unknown';

export interface LegacyOfficeIdentification {
	format: LegacyOfficeFormat;
	support: LegacyOfficeSupport;
	streams: string[];
	reason: string;
}

/** Identify a legacy CFB document from its defining stream names. This detects containers only; it does not parse or claim support for every format. */
export function identifyLegacyOffice(input: Uint8Array | ArrayBuffer): LegacyOfficeIdentification {
	const bytes = input instanceof Uint8Array ? input : new Uint8Array(input);
	try {
		const buffer = bytes.buffer.slice(
			bytes.byteOffset,
			bytes.byteOffset + bytes.byteLength,
		) as ArrayBuffer;
		const streams = parseOle2(buffer)
			.entries.filter((entry) => entry.type === ENTRY_TYPE_STREAM)
			.map((entry) => entry.name);
		const names = new Set(streams.map((name) => name.toLowerCase()));
		if (names.has('worddocument'))
			return result(
				'doc',
				'partial',
				streams,
				'Legacy Word DOC detected; only main-body paragraph text and constrained edits are supported.',
			);
		if (names.has('workbook') || names.has('book'))
			return result(
				'xls',
				'partial',
				streams,
				'Legacy Excel XLS detected; only BIFF8 preview and bounded cell edits are supported.',
			);
		if (names.has('powerpoint document'))
			return result(
				'ppt',
				'partial',
				streams,
				'Legacy PowerPoint PPT detected; record traversal and binary export from the neutral writer model are available. Presentation rendering and viewer-model conversion belong to the PowerPoint viewer.',
			);
		if (names.has('visiodocument')) {
			const inspection = inspectLegacyVisio(bytes);
			return result(
				'visio',
				inspection ? 'partial' : 'unsupported',
				streams,
				inspection
					? 'Visio binary header and trailer inspected; supported version 11 drawing models and guarded text/transform edits are available through parseVsd after separate checked decoding. Earlier versions retain structural inspection only.'
					: 'VisioDocument stream detected, but the binary header is unsupported or invalid.',
			);
		}
		const publication = inspectLegacyPublisher(bytes);
		if (publication)
			return result(
				'publisher',
				'partial',
				streams,
				`${publication.reason} Standard document properties can be edited when present.`,
			);
		if (names.has('contents') && names.has('escher'))
			return result(
				'publisher',
				'unsupported',
				streams,
				'Likely legacy Publisher document detected from its Contents and Escher streams; Publisher binary files are unsupported.',
			);
		return result(
			'unknown',
			'unknown',
			streams,
			'CFB container detected, but its streams do not identify a supported legacy Office format.',
		);
	} catch {
		return result('unknown', 'unknown', [], 'Input is not a readable CFB compound file.');
	}
}

function result(
	format: LegacyOfficeFormat,
	support: LegacyOfficeSupport,
	streams: string[],
	reason: string,
): LegacyOfficeIdentification {
	return { format, support, streams, reason };
}
