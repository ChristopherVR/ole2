import { mkdirSync, readFileSync, writeFileSync } from 'node:fs';
import { resolve, join } from 'node:path';
import { parseXls } from '../dist/index.js';

const directory = process.argv[2];
if (!directory) throw new Error('Usage: node scripts/prepare-native-xls-string-scalar-edits.mjs OUTPUT_DIRECTORY');
const output = resolve(directory);
mkdirSync(output, { recursive: true });
const input = readFileSync(new URL('../test/fixtures/xls/workbook-rich-strings.xls', import.meta.url));
const document = parseXls(input);
const edits = [
  { sheet: 0, row: 0, col: 0, value: 37.125 },
  { sheet: 0, row: 1, col: 0, value: true },
  { sheet: 0, row: 3, col: 0, value: 8.5 },
  { sheet: 0, row: 4, col: 0, value: { error: '#N/A' } },
  { sheet: 1, row: 0, col: 0, value: false },
];
for (const edit of edits) document.sheets[edit.sheet].cell(edit.row, edit.col).value = edit.value;
writeFileSync(join(output, 'original.xls'), input);
writeFileSync(join(output, 'edited.xls'), document.serialize());
writeFileSync(join(output, 'edits.json'), JSON.stringify(edits, null, 2) + '\n');
console.log(output);
