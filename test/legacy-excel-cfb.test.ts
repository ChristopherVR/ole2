import { describe, expect, it } from 'vitest';
import { unwrapXlsBytes } from '../src/legacy-excel-cfb.js';
import { buildOle2 } from '../src/ole2-parser-write.js';
import { readCompoundFileStream } from '../src/ole2-stream-edit.js';
import { nestRootStream } from './helpers/nested-cfb.js';

describe('legacy Excel CFB rewrap safety', () => {
	it('edits a root Workbook in place while preserving sibling streams and sectors', () => {
		const workbook = new Uint8Array([10, 20, 30, 40]);
		const metadata = new Uint8Array([4, 3, 2, 1]);
		const flat = new Uint8Array(
			buildOle2(
				new Map([
					['Workbook', workbook],
					['\u0005SummaryInformation', metadata],
				]),
			),
		);
		const original = flat;
		const unwrap = unwrapXlsBytes(original);
		expect(unwrap.workbookBytes).toEqual(workbook);
		const updated = unwrap.rewrap!(new Uint8Array([10, 20, 30, 99]));
		expect(updated).not.toBe(original);
		expect(readCompoundFileStream(updated, ['Workbook'])).toEqual(new Uint8Array([10, 20, 30, 99]));
		expect(readCompoundFileStream(updated, ['\u0005SummaryInformation'])).toEqual(metadata);
		const unchanged = readCompoundFileStream(updated, ['\u0005SummaryInformation'])!;
		expect(unchanged).toEqual(metadata);
	});

	it('keeps exact bytes when an unsupported resize would flatten a nested storage', () => {
		const flat = new Uint8Array(
			buildOle2(
				new Map([
					['Workbook', new Uint8Array([1, 2, 3])],
					['Other', new Uint8Array([7])],
				]),
			),
		);
		const original = nestRootStream(flat, 'Other', 'ObjectPool');
		const unwrap = unwrapXlsBytes(original);
		expect(unwrap.rewrap!(new Uint8Array([4, 5, 6, 7]))).toBe(original);
	});

	it('keeps the existing conservative behavior for duplicate stream names on resize', () => {
		const flat = new Uint8Array(
			buildOle2(
				new Map([
					['Workbook', new Uint8Array([1, 2, 3])],
					['workboox', new Uint8Array([4, 5, 6])],
				]),
			),
		);
		// Create the malformed collision independently of the validating writer.
		const view = new DataView(flat.buffer);
		const directory = (view.getUint32(0x30, true) + 1) * 512;
		view.setUint16(directory + 256 + 14, 'k'.charCodeAt(0), true);
		const unwrap = unwrapXlsBytes(flat);
		expect(unwrap.workbookBytes).toEqual(flat);
		expect(unwrap.rewrap).toBeUndefined();
	});
});
