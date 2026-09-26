import { describe, expect, it } from 'vitest';
import { readSummaryProperties, writeSummaryTextProperty } from '../src/ole-summary-properties.js';
import {
	readLegacyOfficeMetadata,
	writeLegacyOfficeMetadata,
} from '../src/legacy-office-metadata.js';
import { buildOle2 } from '../src/ole2-parser-write.js';
import { parseOle2 } from '../src/ole2-parser-read.js';

function fixture(type = 30, codePage = 1252): Uint8Array {
	const bytes = new Uint8Array(128),
		view = new DataView(bytes.buffer);
	view.setUint16(0, 0xfffe, true);
	view.setUint32(24, 1, true);
	bytes.set(
		[0xe0, 0x85, 0x9f, 0xf2, 0xf9, 0x4f, 0x68, 0x10, 0xab, 0x91, 8, 0, 0x2b, 0x27, 0xb3, 0xd9],
		28,
	);
	view.setUint32(44, 48, true);
	view.setUint32(48, 80, true);
	view.setUint32(52, 3, true);
	for (const [index, id, offset] of [
		[0, 1, 32],
		[1, 2, 40],
		[2, 14, 72],
	]) {
		view.setUint32(56 + index! * 8, id!, true);
		view.setUint32(60 + index! * 8, offset!, true);
	}
	view.setUint32(80, 2, true);
	view.setUint16(84, codePage, true);
	view.setUint32(88, type, true);
	const unicode = type === 31 || codePage === 1200;
	view.setUint32(92, type === 30 && unicode ? 8 : 4, true);
	for (let i = 0; i < 3; i++) bytes[96 + i * (unicode ? 2 : 1)] = 'Old'.charCodeAt(i);
	view.setUint32(120, 3, true);
	view.setInt32(124, 99, true);
	return bytes;
}

describe('SummaryInformation metadata', () => {
	it.each([0, 4096])(
		'edits metadata through CFB with cutoff %i and preserves other streams',
		(cutoff) => {
			const application = new Uint8Array(5000).fill(42);
			const source = new Uint8Array(
				buildOle2(
					new Map([
						['\u0005SummaryInformation', fixture()],
						['VisioDocument', application],
					]),
					undefined,
					cutoff,
				),
			);
			const updated = writeLegacyOfficeMetadata(source, 'title', 'New title');
			expect(updated).not.toBe(source);
			expect(readLegacyOfficeMetadata(updated)?.title).toBe('New title');
			expect(readLegacyOfficeMetadata(source)?.title).toBe('Old');
			expect(parseOle2(updated.buffer as ArrayBuffer).getStream('VisioDocument')).toEqual(
				application,
			);
			expect(updated.length).toBe(source.length);
		},
	);
	it('edits a property without changing neighboring values, offsets, or input', () => {
		const source = fixture(),
			before = source.slice();
		const updated = writeSummaryTextProperty(source, 'title', 'Euro \u20ac');
		expect(readSummaryProperties(updated)).toEqual({ title: 'Euro \u20ac' });
		expect(source).toEqual(before);
		expect(updated.length).toBe(source.length);
		expect(updated.subarray(0, 92)).toEqual(source.subarray(0, 92));
		expect(updated.subarray(120)).toEqual(source.subarray(120));
	});
	it.each([
		[31, 1200],
		[30, 1200],
	])('preserves UTF16 type %i and length-unit semantics', (type, cp) => {
		const updated = writeSummaryTextProperty(
			fixture(type, cp),
			'title',
			'\u65e5\u672c\ud83d\ude00',
		);
		expect(readSummaryProperties(updated)?.title).toBe('\u65e5\u672c\ud83d\ude00');
		expect(new DataView(updated.buffer).getUint32(92, true)).toBe(type === 31 ? 5 : 10);
	});
	it('returns original bytes for unsupported or unsafe edits', () => {
		const source = fixture();
		for (const value of ['x'.repeat(50), '\u65e5', 'bad\0value', 'Old'])
			expect(writeSummaryTextProperty(source, 'title', value)).toBe(source);
		expect(writeSummaryTextProperty(source, 'author', 'New')).toBe(source);
		const unsupported = fixture(30, 932);
		expect(readSummaryProperties(unsupported)).toEqual({});
		expect(writeSummaryTextProperty(unsupported, 'title', 'New')).toBe(unsupported);
	});
	it('preserves Windows-1250 metadata used by real Visio files', () => {
		const updated = writeSummaryTextProperty(fixture(30, 1250), 'title', '\u0141\u00f3d\u017a');
		expect(readSummaryProperties(updated)?.title).toBe('\u0141\u00f3d\u017a');
		expect(new DataView(updated.buffer).getUint16(84, true)).toBe(1250);
	});
	it('rejects overlapping slots, duplicate properties, wrong FMTID and truncated values', () => {
		for (const [offset, value] of [
			[44, 4],
			[48, 1000],
			[52, 1000],
			[60, 40],
			[64, 1],
			[92, 1000],
			[28, 0],
		]) {
			const malformed = fixture();
			new DataView(malformed.buffer).setUint32(offset!, value!, true);
			expect(readSummaryProperties(malformed)?.title).toBeUndefined();
			expect(writeSummaryTextProperty(malformed, 'title', 'New')).toBe(malformed);
		}
		for (let length = 0; length < 128; length++)
			expect(() => readSummaryProperties(fixture().subarray(0, length))).not.toThrow();
	});
});
