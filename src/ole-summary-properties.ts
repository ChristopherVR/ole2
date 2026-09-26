import { decodeCp1252, encodeCp1252 } from './ole-document-doc-cp1252.js';

/** MS-OLEPS SummaryInformation, shared by legacy Office applications.
 * https://learn.microsoft.com/en-us/openspecs/windows_protocols/ms-oleps/f7933d28-2cc4-4b36-bc23-8861cbcd37c4
 */
export const SUMMARY_STREAM_NAME = '\u0005SummaryInformation';
export const SUMMARY_TEXT_PROPERTIES = {
	title: 2,
	subject: 3,
	author: 4,
	keywords: 5,
	comments: 6,
	template: 7,
	lastAuthor: 8,
	revision: 9,
	application: 18,
} as const;
export type SummaryTextProperty = keyof typeof SUMMARY_TEXT_PROPERTIES;
export type SummaryProperties = Partial<Record<SummaryTextProperty, string>>;
const FMTID = [
	0xe0, 0x85, 0x9f, 0xf2, 0xf9, 0x4f, 0x68, 0x10, 0xab, 0x91, 0x08, 0, 0x2b, 0x27, 0xb3, 0xd9,
];
interface Slot {
	offset: number;
	end: number;
	type: number;
}
interface Layout {
	view: DataView;
	slots: Map<number, Slot>;
	codePage: number;
}
function decoder(codePage: number): TextDecoder | undefined {
	if (codePage < 1250 || codePage > 1258) return;
	try {
		return new TextDecoder(`windows-${codePage}`, { fatal: true });
	} catch {
		return;
	}
}
function encodeText(text: string, codePage: number): Uint8Array | undefined {
	if (codePage === 1252) return encodeCp1252(text);
	const codec = decoder(codePage);
	if (!codec) return;
	const reverse = new Map<string, number>();
	for (let byte = 0; byte < 256; byte++) {
		try {
			reverse.set(codec.decode(Uint8Array.of(byte)), byte);
		} catch {
			/* Unassigned character. */
		}
	}
	const result: number[] = [];
	for (const char of text) {
		const byte = reverse.get(char);
		if (byte === undefined) return;
		result.push(byte);
	}
	return Uint8Array.from(result);
}

function layout(bytes: Uint8Array): Layout | undefined {
	if (bytes.length < 48 || bytes.length > 2_097_152) return;
	const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
	if (
		view.getUint16(0, true) !== 0xfffe ||
		view.getUint16(2, true) > 1 ||
		view.getUint32(24, true) !== 1
	)
		return;
	if (!FMTID.every((value, index) => bytes[28 + index] === value)) return;
	const start = view.getUint32(44, true);
	if (start < 48 || start % 4 || start + 8 > bytes.length) return;
	const size = view.getUint32(start, true),
		count = view.getUint32(start + 4, true);
	if (size < 8 || start + size > bytes.length || count > (size - 8) / 8) return;
	const offsets = new Map<number, number>();
	const seen = new Set<number>();
	for (let i = 0; i < count; i++) {
		const id = view.getUint32(start + 8 + i * 8, true);
		const relative = view.getUint32(start + 12 + i * 8, true);
		if (
			offsets.has(id) ||
			seen.has(relative) ||
			relative % 4 ||
			relative < 8 + count * 8 ||
			relative + 4 > size
		)
			return;
		offsets.set(id, start + relative);
		seen.add(relative);
	}
	const ordered = [...offsets.values()].sort((a, b) => a - b);
	const slots = new Map<number, Slot>();
	for (const [id, offset] of offsets) {
		const end = ordered[ordered.indexOf(offset) + 1] ?? start + size;
		slots.set(id, { offset, end, type: view.getUint32(offset, true) });
	}
	const cp = slots.get(1);
	if (!cp || cp.type !== 2 || cp.offset + 6 > cp.end) return;
	return { view, slots, codePage: view.getUint16(cp.offset + 4, true) };
}

function textSlot(
	bytes: Uint8Array,
	parsed: Layout,
	slot: Slot,
): { text: string; unicode: boolean } | undefined {
	if ((slot.type !== 30 && slot.type !== 31) || slot.offset + 8 > slot.end) return;
	const unicode = slot.type === 31 || parsed.codePage === 1200;
	if (!unicode && !decoder(parsed.codePage)) return;
	const count = parsed.view.getUint32(slot.offset + 4, true);
	const length = slot.type === 31 ? count * 2 : count;
	if (length < (unicode ? 2 : 1) || slot.offset + 8 + length > slot.end || (unicode && length % 2))
		return;
	const data = bytes.subarray(slot.offset + 8, slot.offset + 8 + length);
	if (data[data.length - 1] !== 0 || (unicode && data[data.length - 2] !== 0)) return;
	let text = '';
	if (unicode) {
		for (let i = 0; i < data.length - 2; i += 2)
			text += String.fromCharCode(data[i]! | (data[i + 1]! << 8));
	} else {
		try {
			text =
				parsed.codePage === 1252
					? decodeCp1252(data.subarray(0, -1))
					: decoder(parsed.codePage)!.decode(data.subarray(0, -1));
		} catch {
			return;
		}
	}
	return { text: text.replace(/\0+$/, ''), unicode };
}

/** Read supported text metadata from a raw SummaryInformation stream. Unsupported code pages are omitted. */
export function readSummaryProperties(bytes: Uint8Array): SummaryProperties | undefined {
	const parsed = layout(bytes);
	if (!parsed) return;
	const result: SummaryProperties = {};
	for (const [key, id] of Object.entries(SUMMARY_TEXT_PROPERTIES)) {
		const slot = parsed.slots.get(id);
		const value = slot && textSlot(bytes, parsed, slot);
		if (value) result[key as SummaryTextProperty] = value.text;
	}
	return result;
}

/** Edit an existing text property within its allocated slot. Never moves other properties or changes stream length.
 * Returns original bytes for unsupported, malformed, unrepresentable, or oversized edits.
 */
export function writeSummaryTextProperty(
	bytes: Uint8Array,
	key: SummaryTextProperty,
	value: string,
): Uint8Array {
	if (value.includes('\0')) return bytes;
	const parsed = layout(bytes),
		id = SUMMARY_TEXT_PROPERTIES[key];
	const slot = parsed?.slots.get(id);
	if (!parsed || !slot) return bytes;
	const current = textSlot(bytes, parsed, slot);
	if (!current || current.text === value) return bytes;
	let encoded: Uint8Array | undefined;
	if (current.unicode) {
		encoded = new Uint8Array((value.length + 1) * 2);
		const view = new DataView(encoded.buffer);
		for (let i = 0; i < value.length; i++) view.setUint16(i * 2, value.charCodeAt(i), true);
	} else encoded = encodeText(value + '\0', parsed.codePage);
	if (!encoded || encoded.length > slot.end - slot.offset - 8) return bytes;
	const result = bytes.slice();
	result.fill(0, slot.offset + 8, slot.end);
	result.set(encoded, slot.offset + 8);
	new DataView(result.buffer).setUint32(
		slot.offset + 4,
		slot.type === 31 ? encoded.length / 2 : encoded.length,
		true,
	);
	return result;
}
