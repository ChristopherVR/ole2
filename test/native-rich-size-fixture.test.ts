import { readFileSync } from 'node:fs';
import { expect, it } from 'vitest';
import { readOleDocParagraphs } from '../src/ole-document-doc-editor.js';
import { parseOle2 } from '../src/ole2-parser-read.js';

it('retains the owned direct-size fixture text and plain-story editing scope', () => {
  const bytes = new Uint8Array(readFileSync(new URL('./fixtures/doc/rich-size-runs.doc', import.meta.url)));
  expect(readOleDocParagraphs(bytes)).toEqual(['First plain paragraph.', 'Bold text, italic text, plain text.', 'Last plain paragraph.']);
  const word = parseOle2(bytes.buffer as ArrayBuffer).getStream('WordDocument')!;
  const view = new DataView(word.buffer, word.byteOffset, word.byteLength);
  expect([0x50, 0x54, 0x58, 0x5c, 0x60, 0x64, 0x68].map(offset => view.getUint32(offset, true))).toEqual([0, 0, 0, 0, 0, 0, 0]);
  // Font-size codec regressions live with the guarded writer; independent Word
  // character-font snapshots establish the captured appearance separately.
});
