import assert from 'node:assert/strict';
import { createHash } from 'node:crypto';
import { mkdirSync, readFileSync, writeFileSync } from 'node:fs';
import { join, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import { parseXls } from '../dist/index.js';

// Generate the six previously native-tested candidates; this script never launches Excel.
// Usage after build: node scripts/prepare-native-xls-boolerr-edits.mjs [OUTPUT]
const source = fileURLToPath(new URL('../test/fixtures/xls/workbook-cell-types.xls', import.meta.url));
const output = resolve(process.argv[2] ?? '.native-validation/boolerr-generated');
const hash = bytes => createHash('sha256').update(bytes).digest('hex');
const sourceSha256 = '6cf446bec16da6acca52b48c18521f53fef609e0a2483dd8437a7b9097a89387';
const input = new Uint8Array(readFileSync(source));
assert.equal(hash(input), sourceSha256, 'Owned native fixture must match the pinned source');

// Escapes make the recorded Unicode value independent of terminal encoding.
const text = 'Converted \u6f22\u5b57 \u{1f600}';
const textCodePoints = Array.from(text, character => character.codePointAt(0));
assert.deepEqual(textCodePoints, [67, 111, 110, 118, 101, 114, 116, 101, 100, 32, 28450, 23383, 32, 128512]);
const expectedHashes = [
  ['3b3529ad991aab83d9574fe11ba13a39b726c277444934e18c8313132af053c9', 'bf38206b4af0dbee6c99b9b7d92a9338ba3df81031d268f7c83aa36e3ebe49d6'],
  ['1f4635b8c6131a39c955bf46275fb508590eda64a41442a160f0b36f8e44bbe1', '54149ce360fa984ed45363c4331e7e44f8d131c3da0e93d0c0bdf387ba022d96'],
  ['05b259ee103fd909aec78e015586c5cc117982e5b9cc0606fbe3481038f203f1', '116be664955497ab031602056ca775c77ce2ad273ddb758d229ef9480df65295'],
];

function assertAtomicRefusal(document, action) {
  const before = document.serialize();
  const revision = document.revision;
  const dirty = document.dirty;
  const workbook = document.workbook;
  assert.throws(action);
  assert.deepEqual(document.serialize(), before);
  assert.equal(document.revision, revision);
  assert.equal(document.dirty, dirty);
  assert.deepEqual(document.workbook, workbook);
}

// Prove refusal before and after a successful edit leaves the entire model unchanged.
const refused = parseXls(input);
assertAtomicRefusal(refused, () => { refused.sheets[0].cell(0, 1).value = Infinity; });
refused.sheets[0].cell(0, 1).value = 37.125;
assertAtomicRefusal(refused, () => { refused.sheets[0].cell(0, 1).value = null; });

const cases = [];
// Fresh source for each case: B1=false, C1=#DIV/0!, D1=#N/A.
for (let col = 1; col <= 3; col++) {
  for (const [kind, value, hashIndex] of [['number', 37.125, 0], ['string', text, 1]]) {
    const document = parseXls(input);
    const cell = document.sheets[0].cell(0, col);
    const xf = cell.xf;
    assert.equal(xf, 15, 'Expected source BOOLERR cell formatting');
    cell.value = value;
    assert.equal(document.dirty, true);
    assert.equal(document.revision, 1);
    const bytes = document.serialize();
    const sha256 = hash(bytes);
    assert.equal(sha256, expectedHashes[col - 1][hashIndex], 'Must reproduce the native-tested bytes');
    const reread = parseXls(bytes).sheets[0].cell(0, col);
    assert.equal(reread.value, value);
    assert.equal(reread.xf, xf);
    cases.push({ name: `boolerr-${col}-${kind}`, sheet: 0, row: 0, col, value,
      valueCodePoints: kind === 'string' ? textCodePoints : undefined,
      sha256, byteLength: bytes.length, bytes });
  }
}

// Validate all candidates before writing any output.
mkdirSync(output, { recursive: true });
const records = cases.map(({ bytes, ...record }) => {
  const path = join(output, `${record.name}.xls`);
  writeFileSync(path, bytes);
  return { ...record, path };
});
const manifest = {
  source, sourceSha256, generatorSha256: hash(readFileSync(new URL(import.meta.url))),
  nativeEvidenceWriter: '1aa1dfb063c6923df470285f53d898a4b249aba7',
  nativeEvidenceSha256: 'd83335b5ab0f38c93a219c89181fe49fe9cbf414fe03fdf2112efabb95ea4665',
  nativeValidatedByThisRun: false, fullFidelity: false,
  atomicRefusalAssertions: ['non-finite number before edit', 'clearing to null after edit'],
  cases: records,
};
writeFileSync(join(output, 'manifest.json'), JSON.stringify(manifest, null, 2));
console.log(JSON.stringify({ generated: records.length, output, nativeValidatedByThisRun: false }));
