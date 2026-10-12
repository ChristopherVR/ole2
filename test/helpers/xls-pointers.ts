import { expect } from 'vitest';
import { unwrapXlsBytes } from '../../src/legacy-excel-cfb.js';
import { readRecords } from '../../src/legacy-excel-biff8.js';

/** Validate BIFF worksheet/row-block addresses independently of cell reading. */
export function checkXlsPointers(input: Uint8Array): void {
	const bytes = unwrapXlsBytes(input).workbookBytes, view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
	const records = readRecords(bytes, 0, bytes.length), byOffset = new Map(records.map(r => [r.headerOffset, r]));
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
