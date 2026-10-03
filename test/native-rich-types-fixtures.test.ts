import { readFileSync } from 'node:fs';
import { expect, it } from 'vitest';
import { readOleDocParagraphs } from '../src/ole-document-doc-editor.js';
import { parseOle2 } from '../src/ole2-parser-read.js';
import { readRecords } from '../src/legacy-excel-biff8.js';

const fixture = (path: string) => new Uint8Array(readFileSync(new URL(`./fixtures/${path}`, import.meta.url)));
it('native rich fixture has exactly three plain paragraphs and no other populated stories', () => {
  const bytes = fixture('doc/rich-runs.doc');
  expect(readOleDocParagraphs(bytes)).toEqual(['First plain paragraph.', 'Bold text, italic text, plain text.', 'Last plain paragraph.']);
  const word = parseOle2(bytes.buffer as ArrayBuffer).getStream('WordDocument')!;
  const view = new DataView(word.buffer, word.byteOffset, word.byteLength);
  expect([0x50, 0x54, 0x58, 0x5c, 0x60, 0x64, 0x68].map(offset => view.getUint32(offset, true))).toEqual([0, 0, 0, 0, 0, 0, 0]);
});
it('native cell fixture stores true/false/errors as four actual BOOLERR records', () => {
  const bytes = fixture('xls/workbook-cell-types.xls');
  const stream = parseOle2(bytes.buffer as ArrayBuffer).getStream('Workbook')!;
  const view = new DataView(stream.buffer, stream.byteOffset, stream.byteLength);
  const values = readRecords(stream, 0, stream.length).filter(record => record.opcode === 0x0205).map(record => ({
    row: view.getUint16(record.dataOffset, true), col: view.getUint16(record.dataOffset + 2, true),
    value: stream[record.dataOffset + 6], error: stream[record.dataOffset + 7],
  }));
  expect(values).toEqual([
    { row: 0, col: 0, value: 1, error: 0 }, { row: 0, col: 1, value: 0, error: 0 },
    { row: 0, col: 2, value: 7, error: 1 }, { row: 0, col: 3, value: 42, error: 1 },
  ]);
});
