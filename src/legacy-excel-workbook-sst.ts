/**
 * The BIFF8 shared string table (`SST`) with its `CONTINUE` records stitched.
 *
 * Each entry is an `XLUnicodeRichExtendedString`: `cch:u16`, flags
 * (`fHighByte`, `fExtSt`, `fRichSt`), optional run count and phonetic block
 * size, the characters, then the formatting runs and phonetic data, which
 * this reader skips (rich runs and phonetic guides are not modelled). A
 * string split across a `CONTINUE` boundary restarts with a one-byte
 * high-byte flag that can switch between compressed and UTF-16 characters.
 *
 * @module legacy-excel-workbook-sst
 */
import { SegmentReader, type XlsRecord } from './legacy-excel-workbook-records.js';

/** Decode every string of an `SST` record and its continuations. */
export function parseSharedStrings(record: XlsRecord): string[] {
	const reader = SegmentReader.of(record);
	reader.u32(); // cstTotal: references in the workbook
	const unique = reader.u32();
	const strings: string[] = [];
	// Some writers leave cstUnique at 0; read to the end of the record in that case.
	while ((unique === 0 || strings.length < unique) && !reader.done) {
		const count = reader.u16();
		const flags = reader.u8();
		const runs = flags & 0x08 ? reader.u16() : 0;
		const extSize = flags & 0x04 ? reader.u32() : 0;
		strings.push(reader.continuedChars(count, (flags & 1) !== 0));
		reader.skip(runs * 4 + extSize);
	}
	return strings;
}
