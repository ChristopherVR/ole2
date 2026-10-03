import { readFileSync } from 'node:fs';
import { describe, expect, it } from 'vitest';
import { readOleDocParagraphs } from '../src/ole-document-doc-editor.js';
import { parseOle2 } from '../src/ole2-parser-read.js';
import { readRecords } from '../src/legacy-excel-biff8.js';

const fixture = (name: string) => new Uint8Array(readFileSync(new URL(`./fixtures/${name}`, import.meta.url)));
describe('native authored conversion corpus', () => {
  it('contains one main-story field between plain paragraphs and no populated other stories', () => {
    const bytes = fixture('doc/main-field.doc');
    expect(readOleDocParagraphs(bytes)).toEqual([
      'First paragraph plain text.',
      'Second paragraph has a field: \u0013 QUOTE "Fixture field" \u0014Fixture field\u0015',
      'Third paragraph plain text.',
    ]);
    const word = parseOle2(bytes.buffer as ArrayBuffer).getStream('WordDocument')!;
    const view = new DataView(word.buffer, word.byteOffset, word.byteLength);
    expect([0x50, 0x54, 0x58, 0x5c, 0x60, 0x64, 0x68].map(offset => view.getUint32(offset, true))).toEqual([0, 0, 0, 0, 0, 0, 0]);
  });
  it('contains a real four-cell MULRK whose entries have different XF styles', () => {
    const bytes = fixture('xls/workbook-mulrk.xls');
    const workbook = parseOle2(bytes.buffer as ArrayBuffer).getStream('Workbook')!;
    const view = new DataView(workbook.buffer, workbook.byteOffset, workbook.byteLength);
    const mulrk = readRecords(workbook, 0, workbook.length).find(record => record.opcode === 0x00bd && view.getUint16(record.dataOffset, true) === 0)!;
    expect(mulrk).toBeDefined();
    expect(view.getUint16(mulrk.dataOffset + 2, true)).toBe(0);
    expect(view.getUint16(mulrk.dataOffset + mulrk.length - 2, true)).toBe(3);
    const xfs = Array.from({ length: 4 }, (_, index) => view.getUint16(mulrk.dataOffset + 4 + index * 6, true));
    expect(new Set(xfs).size).toBe(4);
  });
});
