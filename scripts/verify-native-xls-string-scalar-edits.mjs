import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import { join } from 'node:path';

const directory = process.argv[2];
if (!directory) throw new Error('Usage: node scripts/verify-native-xls-string-scalar-edits.mjs SNAPSHOT_DIRECTORY');
const read = name => JSON.parse(readFileSync(join(directory, name), 'utf8').replace(/^\uFEFF/, ''));
const original = read('original.json'), edits = read('edits.json');
const selected = (sheet, cell) => edits.some(edit => edit.sheet === sheet && edit.row === cell.row && edit.col === cell.col);
const unrelated = snapshot => ({ ...snapshot, sheets: snapshot.sheets.map((sheet, index) => ({
  ...sheet, cells: sheet.cells.filter(cell => !selected(index, cell)),
})) });
for (const name of ['edited.json', 'reopened.json']) {
  const snapshot = read(name);
  assert.equal(snapshot.consumer, 'Microsoft Excel');
  assert.equal(snapshot.corruptLoad, 0);
  assert.notEqual(snapshot.repairMode, true);
  assert.deepEqual(unrelated(snapshot), unrelated(original), 'Unrelated captured fields must match');
  for (const edit of edits) {
    const cell = snapshot.sheets[edit.sheet].cells.find(cell => cell.row === edit.row && cell.col === edit.col);
    assert(cell, 'Edited cell must exist');
    if (typeof edit.value === 'object') {
      assert.equal(cell.isError, true); assert.equal(cell.displayText, edit.value.error);
    } else {
      assert.equal(cell.isError, false); assert.equal(cell.value, edit.value);
    }
    const before = original.sheets[edit.sheet].cells.find(cell => cell.row === edit.row && cell.col === edit.col);
    assert.equal(cell.numberFormat, before.numberFormat); assert.equal(cell.merge, before.merge);
  }
}
console.log('Native Excel string-to-scalar open/save/reopen passed for five declared edits; unrelated captured fields match.');
