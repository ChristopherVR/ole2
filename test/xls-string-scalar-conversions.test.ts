import { readFileSync } from 'node:fs';
import { describe, expect, it } from 'vitest';
import { parseXls, readXlsWorkbook, editXlsStringScalarCell, editXlsStringScalarWorkbookStream,
	buildCompoundFile, listCompoundFile, type XlsStringScalarEdit } from '../src/index.js';
import { unwrapXlsBytes } from '../src/legacy-excel-cfb.js';
import { readRecords } from '../src/legacy-excel-biff8.js';
import { readPreservedXlsSst } from '../src/legacy-excel-preserved-sst.js';
import { checkXlsPointers } from './helpers/xls-pointers.js';

const fixture = () => new Uint8Array(readFileSync(new URL('./fixtures/xls/workbook-rich-strings.xls', import.meta.url)));
const raw = (bytes: Uint8Array) => unwrapXlsBytes(bytes).workbookBytes;
const values: XlsStringScalarEdit['value'][] = [37.125, -0, Number.MIN_VALUE, true, false, { error: '#N/A' }, { error: '#REF!' }];

describe('existing XLS string to scalar conversions', () => {
	it.each(values.map((value, index) => ({ value, index })))('converts plain shared strings in scalar case $index without changing aliases', ({ value }) => {
		const input = fixture(), before = readXlsWorkbook(input), source = raw(input), oldSst = readPreservedXlsSst(source);
		const document = parseXls(input), cell = document.sheets[0]!.cell(0, 0), alias = document.sheets[0]!.cell(0, 1);
		const aliasValue = alias.value, xf = cell.xf;
		cell.value = value;
		expect(cell.value).toEqual(value); expect(cell.xf).toBe(xf); expect(alias.value).toEqual(aliasValue);
		expect(document.revision).toBe(1); expect(document.recalculationRequired).toBe(true);
		const saved = document.serialize(), after = readXlsWorkbook(saved);
		checkXlsPointers(saved);
		const selected = after.sheets[0]!.cells.find(c => c.row === 0 && c.col === 0)!;
		selected.value = before.sheets[0]!.cells.find(c => c.row === 0 && c.col === 0)!.value;
		expect(after).toEqual(before);
		const next = raw(saved), newSst = readPreservedXlsSst(next);
		expect(newSst.total).toBe(oldSst.total - 1);
		const normalizedSst = next.slice(newSst.start, newSst.end);
		new DataView(normalizedSst.buffer).setUint32(4, oldSst.total, true);
		expect(normalizedSst).toEqual(source.slice(oldSst.start, oldSst.end));
		for (const node of listCompoundFile(input).nodes.filter(n => n.path[0] !== 'Workbook'))
			expect(listCompoundFile(saved).nodes.find(n => JSON.stringify(n.path) === JSON.stringify(node.path))).toEqual(node);
		const ops = new Set([0x6, 0x201, 0x203, 0x204, 0x205, 0x27e, 0xfd, 0xbd, 0xbe]);
		const unchangedCells = (bytes: Uint8Array) => {
			const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
			let sheet = -1;
			return readRecords(bytes, 0, bytes.length).filter(r => {
				if (r.opcode === 0x809) sheet++;
				return ops.has(r.opcode) && !(sheet === 1 && view.getUint16(r.dataOffset, true) === 0 && view.getUint16(r.dataOffset + 2, true) === 0);
			}).map(r => bytes.slice(r.headerOffset, r.dataOffset + r.length));
		};
		expect(unchangedCells(next)).toEqual(unchangedCells(source));
		expect(parseXls(saved).sheets[0]!.cell(0, 0).value).toEqual(value);
	});

	it.each([1, 2, 3, 4])('converts rich or continued strings on row %i while retaining the entire SST payload', row => {
		const input = fixture(), document = parseXls(input), source = raw(input), before = readPreservedXlsSst(source);
		document.sheets[0]!.cell(row, 0).value = 2.75;
		const next = raw(document.serialize()), after = readPreservedXlsSst(next);
		checkXlsPointers(document.serialize());
		expect(next.slice(after.start + 8, after.end)).toEqual(source.slice(before.start + 8, before.end));
		expect(document.sheets[0]!.cell(1, 1).value).toBe('Rich preserved Ω');
		expect(after.total).toBe(before.total - 1);
	});

	it('keeps handles current across repeated string, boolean, error and numeric conversions', () => {
		const document = parseXls(fixture()), cell = document.sheets[0]!.cell(4, 0);
		const count = readPreservedXlsSst(raw(document.serialize())).total;
		for (const value of [true, 3.25, 'Restored Ω', { error: '#DIV/0!' } as const, 'Again', -0]) cell.value = value;
		expect(Object.is(cell.value, -0)).toBe(true); expect(document.revision).toBe(6);
		expect(readPreservedXlsSst(raw(document.serialize())).total).toBe(count - 1);
		checkXlsPointers(document.serialize());
		const baseline = document.serialize(); cell.value = -0;
		expect(document.serialize()).toEqual(baseline); expect(document.revision).toBe(6);
	});

	it('supports later worksheets, v4 containers and Buffer inputs without mutating caller bytes', () => {
		const source = fixture(), listed = listCompoundFile(source);
		for (const input of [source, buildCompoundFile(listed.nodes, { ...listed.options, majorVersion: 4 }), Buffer.from(raw(source))]) {
			const original = new Uint8Array(input);
			const result = editXlsStringScalarCell(input, { row: 0, col: 0, worksheetIndex: 1, value: false });
			expect(result.status).toBe('edited'); expect(new Uint8Array(input)).toEqual(original);
			checkXlsPointers(result.bytes);
			expect(readXlsWorkbook(result.bytes).sheets[1]!.cells.find(c => c.row === 0 && c.col === 0)?.value).toBe(false);
		}
	});

	it('refuses malformed SST counts, references and pointer-bearing records atomically', () => {
		const source = raw(fixture()), sst = readPreservedXlsSst(source), records = readRecords(source, 0, source.length);
		const label = records.find(r => r.opcode === 0xfd)!, db = records.find(r => r.opcode === 0xd7)!;
		for (const mutate of [
			(v: DataView) => v.setUint32(sst.start + 4, 0, true),
			(v: DataView) => v.setUint32(label.dataOffset + 6, sst.entries.length, true),
			(v: DataView) => v.setUint32(db.dataOffset, 1, true),
			(v: DataView) => v.setUint16(records.find(r => r.opcode === 0x55)!.headerOffset, 0x7777, true),
		]) {
			const input = source.slice(); mutate(new DataView(input.buffer)); const original = input.slice();
			for (const value of [8, true]) {
				const result = editXlsStringScalarWorkbookStream(input, { row: 0, col: 0, value });
				expect(result.status).toBe('unchanged'); expect(result.bytes).toBe(input); expect(input).toEqual(original);
			}
		}
	});

	it('refuses formula targets, merged followers, invalid values and clearing without changing document state', () => {
		const input = fixture(), document = parseXls(input);
		const formula = document.sheets[1]!.cell(2, 0); expect(formula.type).toBe('formula');
		expect(() => { formula.value = true; }).toThrow();
		for (const value of [null, NaN, Infinity, { error: '#INVALID!' }])
			expect(() => { document.sheets[0]!.cell(0, 0).value = value as never; }).toThrow();
		const source = raw(input), merged = readRecords(source, 0, source.length).find(r => r.opcode === 0xe5)!;
		const malformed = source.slice(), view = new DataView(malformed.buffer);
		view.setUint16(merged.dataOffset + 2, 0, true); view.setUint16(merged.dataOffset + 4, 0, true);
		view.setUint16(merged.dataOffset + 6, 0, true); view.setUint16(merged.dataOffset + 8, 1, true);
		const refusal = editXlsStringScalarWorkbookStream(malformed, { row: 0, col: 1, value: true });
		expect(refusal.status).toBe('unchanged'); expect(refusal.bytes).toBe(malformed);
		expect(document.serialize()).toEqual(input); expect(document.revision).toBe(0); expect(document.dirty).toBe(false);
	});
});
