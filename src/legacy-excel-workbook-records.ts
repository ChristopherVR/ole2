/**
 * BIFF8 record framing for the workbook reader: records with their trailing
 * `CONTINUE` records attached, a reader that walks a record and its
 * continuations as one logical byte stream, and the BIFF8 string layouts
 * (`XLUnicodeString`, `ShortXLUnicodeString`, continued character runs whose
 * high-byte flag is re-read at every `CONTINUE` boundary).
 *
 * @module legacy-excel-workbook-records
 */

/** BIFF8 `CONTINUE` record opcode. */
export const RT_CONTINUE = 0x003c;

/** One logical BIFF8 record: its own data plus any immediately following `CONTINUE` records. */
export interface XlsRecord {
	opcode: number;
	/** Offset of the record header in the workbook stream (BOUNDSHEET offsets point at these). */
	offset: number;
	data: Uint8Array;
	continues: Uint8Array[];
}

/** Read every record in `[start, end)`, attaching `CONTINUE` records to the record they extend. */
export function readXlsRecords(bytes: Uint8Array, start = 0, end = bytes.length): XlsRecord[] {
	const records: XlsRecord[] = [];
	let offset = start;
	while (offset + 4 <= end) {
		const opcode = bytes[offset]! | (bytes[offset + 1]! << 8);
		const length = bytes[offset + 2]! | (bytes[offset + 3]! << 8);
		const dataStart = offset + 4;
		if (dataStart + length > end) break;
		const data = bytes.subarray(dataStart, dataStart + length);
		const previous = records[records.length - 1];
		if (opcode === RT_CONTINUE && previous) previous.continues.push(data);
		else records.push({ opcode, offset, data, continues: [] });
		offset = dataStart + length;
	}
	return records;
}

/** Little-endian reader over one byte array. Reads past the end yield 0 (never throw). */
export class ByteReader {
	pos: number;
	constructor(
		readonly bytes: Uint8Array,
		pos = 0,
	) {
		this.pos = pos;
	}
	get remaining(): number {
		return this.bytes.length - this.pos;
	}
	u8(): number {
		const value = this.bytes[this.pos] ?? 0;
		this.pos += 1;
		return value;
	}
	u16(): number {
		const value = (this.bytes[this.pos] ?? 0) | ((this.bytes[this.pos + 1] ?? 0) << 8);
		this.pos += 2;
		return value;
	}
	i16(): number {
		const value = this.u16();
		return value >= 0x8000 ? value - 0x10000 : value;
	}
	u32(): number {
		return (this.u16() | (this.u16() << 16)) >>> 0;
	}
	f64(): number {
		const slice = new Uint8Array(8);
		for (let i = 0; i < 8; i++) slice[i] = this.bytes[this.pos + i] ?? 0;
		this.pos += 8;
		return new DataView(slice.buffer).getFloat64(0, true);
	}
	skip(count: number): void {
		this.pos += count;
	}
	/** `count` characters, one byte each when `wide` is false (Latin-1 "compressed" UTF-16). */
	chars(count: number, wide: boolean): string {
		let text = '';
		for (let i = 0; i < count && this.pos < this.bytes.length; i++) {
			text += String.fromCharCode(wide ? this.u16() : this.u8());
		}
		return text;
	}
	/** `XLUnicodeString`: `cch:u16`, `fHighByte` flags byte, characters (no rich/ext data). */
	xlString(): string {
		const count = this.u16();
		const flags = this.u8();
		return this.chars(count, (flags & 1) !== 0);
	}
	/** `ShortXLUnicodeString`: `cch:u8`, flags byte, characters. */
	shortXlString(): string {
		const count = this.u8();
		const flags = this.u8();
		return this.chars(count, (flags & 1) !== 0);
	}
	/** `charCount` UTF-16LE characters (HLINK strings), cut at the first NUL. */
	utf16(charCount: number): string {
		const text = this.chars(charCount, true);
		const nul = text.indexOf('\u0000');
		return nul >= 0 ? text.slice(0, nul) : text;
	}
}

/**
 * A reader over a record's data followed by its `CONTINUE` payloads. Plain
 * reads cross segment boundaries transparently; {@link continuedChars}
 * implements the string rule that a character run split by `CONTINUE`
 * restarts with a new high-byte flag byte.
 */
export class SegmentReader {
	private segment = 0;
	private pos = 0;
	constructor(private readonly segments: Uint8Array[]) {}

	static of(record: XlsRecord): SegmentReader {
		return new SegmentReader([record.data, ...record.continues]);
	}

	get done(): boolean {
		this.normalize();
		return this.segment >= this.segments.length;
	}

	private normalize(): void {
		while (
			this.segment < this.segments.length &&
			this.pos >= (this.segments[this.segment]?.length ?? 0)
		) {
			this.segment++;
			this.pos = 0;
		}
	}

	/** True when the current segment has no bytes left (a character run would need a new flag byte). */
	private atSegmentEnd(): boolean {
		return this.pos >= (this.segments[this.segment]?.length ?? 0);
	}

	u8(): number {
		this.normalize();
		const value = this.segments[this.segment]?.[this.pos] ?? 0;
		this.pos++;
		return value;
	}
	u16(): number {
		return this.u8() | (this.u8() << 8);
	}
	u32(): number {
		return (this.u16() | (this.u16() << 16)) >>> 0;
	}
	skip(count: number): void {
		for (let i = 0; i < count; i++) this.u8();
	}

	/**
	 * Read `count` characters starting with width `wide`. When the current
	 * segment ends before the run does, the next segment begins with a fresh
	 * flags byte whose bit 0 selects the width of the remaining characters.
	 */
	continuedChars(count: number, wide: boolean): string {
		let text = '';
		let isWide = wide;
		for (let read = 0; read < count; read++) {
			if (this.atSegmentEnd()) {
				if (this.segment + 1 >= this.segments.length) break;
				this.segment++;
				this.pos = 0;
				isWide = (this.u8() & 1) !== 0;
			}
			text += String.fromCharCode(isWide ? this.u16() : this.u8());
		}
		return text;
	}

	/** `XLUnicodeString` whose characters may continue into the next segment. */
	xlString(): string {
		const count = this.u16();
		const flags = this.u8();
		return this.continuedChars(count, (flags & 1) !== 0);
	}
}

/** Concatenate a record's data and continuations (for records whose continuations are plain bytes). */
export function joinedData(record: XlsRecord): Uint8Array {
	if (record.continues.length === 0) return record.data;
	const total = record.continues.reduce((sum, part) => sum + part.length, record.data.length);
	const out = new Uint8Array(total);
	out.set(record.data, 0);
	let at = record.data.length;
	for (const part of record.continues) {
		out.set(part, at);
		at += part.length;
	}
	return out;
}
