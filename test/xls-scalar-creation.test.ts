import { readFileSync } from 'node:fs';
import { describe, expect, it } from 'vitest';
import { createXlsCell, createXlsCellWorkbookStream, parseXls, readXlsWorkbook,
	listCompoundFile, buildCompoundFile, type XlsCreatableCellValue } from '../src/index.js';
import { unwrapXlsBytes } from '../src/legacy-excel-cfb.js';
import { readRecords } from '../src/legacy-excel-biff8.js';

const fixture = (name = 'workbook-features.xls') => new Uint8Array(readFileSync(new URL(`./fixtures/xls/${name}`, import.meta.url)));
const raw = (input: Uint8Array) => unwrapXlsBytes(input).workbookBytes;
const values: XlsCreatableCellValue[] = [Math.PI, -0, '', 'Unicode Ω 日本 😀', '日本'.repeat(12000), true, false, { error: '#N/A' }, { error: '#DIV/0!' }];

function checkPointers(input: Uint8Array): void {
	const bytes = raw(input), view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
	const records = readRecords(bytes, 0, bytes.length), byOffset = new Map(records.map(record => [record.headerOffset, record]));
	for (const record of records) {
		if (record.opcode === 0x85) expect(byOffset.get(view.getUint32(record.dataOffset, true))?.opcode).toBe(0x809);
		if (record.opcode === 0x20b) for (let p = record.dataOffset + 12; p < record.dataOffset + record.length; p += 4)
			expect(byOffset.get(view.getUint32(p, true))?.opcode).toBe(p === record.dataOffset + 12 ? 0x55 : 0xd7);
		if (record.opcode !== 0xd7 || !view.getUint32(record.dataOffset, true)) continue;
		const first = byOffset.get(record.headerOffset - view.getUint32(record.dataOffset, true))!;
		expect(first.opcode).toBe(0x208);
		const rows = records.filter(row => row.opcode === 0x208 && row.headerOffset >= first.headerOffset && row.headerOffset < record.headerOffset);
		expect(rows.length).toBe((record.length - 4) / 2);
		let address = first.dataOffset + first.length;
		for (let i = 0; i < rows.length; i++) {
			address += view.getUint16(record.dataOffset + 4 + i * 2, true);
			expect(view.getUint16(byOffset.get(address)!.dataOffset, true)).toBe(view.getUint16(rows[i]!.dataOffset, true));
		}
	}
}

describe('explicit XLS scalar cell creation', () => {
	it.each(values.map((value, index) => ({ value, index })))('creates scalar case $index and preserves existing cells, records and streams', ({ value }) => {
		const input = fixture(), before = readXlsWorkbook(input), beforeStreams = listCompoundFile(input).nodes;
		const document = parseXls(input), existing = document.sheets[0]!.cell(1, 0), existingValue = existing.value;
		const created = document.sheets[0]!.createCell(2, 6, value, 15);
		expect(created.value).toEqual(value); expect(created.xf).toBe(15);
		expect(document.revision).toBe(1); expect(document.dirty).toBe(true); expect(document.recalculationRequired).toBe(true);
		expect(existing.value).toEqual(existingValue);
		const saved = document.serialize(), after = readXlsWorkbook(saved);
		const selected = after.sheets[0]!.cells.find(cell => cell.row === 2 && cell.col === 6)!;
		expect(selected.value).toEqual(value);
		after.sheets[0]!.cells = after.sheets[0]!.cells.filter(cell => cell !== selected);
		expect(after).toEqual(before); checkPointers(saved);
		const afterStreams = listCompoundFile(saved).nodes;
		for (const node of beforeStreams.filter(node => node.path[0] !== 'Workbook'))
			expect(afterStreams.find(other => JSON.stringify(other.path) === JSON.stringify(node.path))).toEqual(node);
		const cellOps = new Set([0x6, 0x201, 0x203, 0x204, 0x205, 0x27e, 0xfd, 0xbd, 0xbe, 0xd6]);
		const cellBytes = (bytes: Uint8Array, exclude = false) => {
			const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
			return readRecords(bytes, 0, bytes.length).filter(record => cellOps.has(record.opcode) &&
				!(exclude && view.getUint16(record.dataOffset, true) === 2 && view.getUint16(record.dataOffset + 2, true) === 6))
				.map(record => bytes.slice(record.headerOffset, record.dataOffset + record.length));
		};
		expect(cellBytes(raw(saved), true)).toEqual(cellBytes(raw(input)));
	});

	it('preserves packed numeric predecessors and supports subsequent edits through stable handles', () => {
		const document = parseXls(fixture('workbook-mulrk.xls')), sheet = document.sheets[0]!;
		const first = sheet.createCell(3, 2, 7, 62);
		const second = sheet.createCell(3, 3, true, 63);
		second.value = { error: '#REF!' }; first.value = 'Updated';
		expect(first.value).toBe('Updated'); expect(second.value).toEqual({ error: '#REF!' });
		checkPointers(document.serialize());
	});

	it('supports whole-v4 and bare Workbook inputs', () => {
		const listed = listCompoundFile(fixture());
		const v4 = buildCompoundFile(listed.nodes, { ...listed.options, majorVersion: 4 });
		for (const input of [v4, raw(fixture())]) {
			const result = createXlsCell(input, { row: 2, col: 6, value: 'Created', xf: 15 });
			expect(result.status).toBe('edited');
			expect(readXlsWorkbook(result.bytes).sheets[0]!.cells.find(cell => cell.row === 2 && cell.col === 6)?.value).toBe('Created');
			checkPointers(result.bytes);
		}
	});

	it('rolls back staged insertion when string conversion rejects a malformed SST', () => {
		const input = raw(fixture()).slice(), records = readRecords(input, 0, input.length);
		const sst = records.find(record => record.opcode === 0xfc)!;
		new DataView(input.buffer, input.byteOffset, input.byteLength).setUint32(sst.dataOffset, 0x7fffffff, true);
		const original = input.slice();
		const result = createXlsCellWorkbookStream(input, { row: 2, col: 6, value: 'Rejected', xf: 15 });
		expect(result.status).toBe('unchanged'); expect(result.bytes).toBe(input); expect(input).toEqual(original);
	});

	it.each([null, undefined, Infinity, NaN, '\ud800', 'x'.repeat(32768), { error: '#INVALID!' }])('rejects invalid value %j without committing', value => {
		const input = fixture(), document = parseXls(input);
		expect(() => document.sheets[0]!.createCell(2, 6, value as XlsCreatableCellValue, 15)).toThrow();
		expect(document.serialize()).toEqual(input); expect(document.revision).toBe(0);
		expect(document.dirty).toBe(false); expect(document.recalculationRequired).toBe(false);
	});

	it('keeps previous edits on refused overwrites and unsupported row/style requests', () => {
		const document = parseXls(fixture()), sheet = document.sheets[0]!;
		const created = sheet.createCell(2, 6, 'Kept', 15), committed = document.serialize();
		for (const action of [() => sheet.createCell(2, 6, true, 15), () => sheet.createCell(65535, 6, true, 15),
			() => sheet.createCell(3, 3, true, 0)]) {
			expect(action).toThrow(); expect(document.serialize()).toEqual(committed); expect(document.revision).toBe(1);
			expect(created.value).toBe('Kept');
		}
	});

	it('captures error values once and refuses accessor-triggered reentrant edits', () => {
		const document = parseXls(fixture()); let reads = 0;
		const value = { get error() { reads++; return '#N/A' as const; } };
		document.sheets[0]!.createCell(2, 6, value, 15); expect(reads).toBe(1);
		const baseline = document.serialize();
		const reentrant = { get error() { document.sheets[0]!.cell(1, 0).value = 'Nested'; return '#N/A' as const; } };
		expect(() => document.sheets[0]!.createCell(3, 3, reentrant, 15)).toThrow();
		expect(document.serialize()).toEqual(baseline); expect(document.revision).toBe(1);
	});
});
