import { readFileSync, writeFileSync, mkdirSync, readdirSync } from 'node:fs';
import { createHash } from 'node:crypto';
import { resolve, dirname, join } from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';
import assert from 'node:assert/strict';

// Entry point: node scripts/prepare-native-vsd-text-matrix.mjs [MODULE] [OUTPUT].
// Candidate generation and atomic-refusal assertions only. Native Visio open,
// captured-field comparisons and save/reopen remain separate validation steps.
const root = resolve(dirname(fileURLToPath(import.meta.url)), '..');
const modulePath = resolve(process.argv[2] ?? join(root, 'dist/index.js'));
const out = resolve(process.argv[3] ?? join(root, '.native-validation/vsd-text-matrix'));
mkdirSync(out, { recursive: true });
const { parseVsd } = await import(pathToFileURL(modulePath));
const sha = bytes => createHash('sha256').update(bytes).digest('hex');

// Fixed xorshift32 seed and draw order retain the native-tested input matrix.
let random = 0xc0deface;
const seed = random;
const next = () => {
  random ^= random << 13;
  random ^= random >>> 17;
  random ^= random << 5;
  return random >>> 0;
};
const draw = (alphabet, n) =>
  Array.from({ length: n }, () => alphabet[next() % alphabet.length]).join('');

const fixtures = {
  literal: [
    'native-literal-transform.vsd',
    '14496febd65e0be62aa8fde493c0f07e812ac6181165b18899b2c1c8ea37bec7',
  ],
  original: [
    'native-visio16-v11.vsd',
    'c6c97822e7bb2cc3e96da9d7d35fe74ac16a7f2c90d19ad4c80967d2896e0974',
  ],
};
const inputs = {};
for (const [k, [name, hash]] of Object.entries(fixtures)) {
  inputs[k] = readFileSync(join(root, 'test/fixtures/vsd', name));
  assert.equal(sha(inputs[k]), hash);
}

const plans = [];
const add = (name, fixture, shapeId, words, saveReopen = false) =>
  plans.push({ name, fixture, shapeId, words, saveReopen });
const bmp = [
  ...Array.from('\u03a9\u03b2\u6f22\u5b57\u6771\u4eac\u0430\u0431\u00e9\u03a3'),
];
const ascii = Array.from('ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789!?#$%');

for (const [name, text] of [
  ['ascii-entropy', 'Q7z!?'],
  ['ascii-repeat', 'AAAAA'],
  ['ascii-seeded', draw(ascii, 5)],
  ['bmp-seeded', draw(bmp, 5)],
  ['bmp-mixed', '\u03a9\u6f22\u5b57\u0430\u0431'],
  ['combining', 'A\u0301B\u0327C'],
  ['emoji', 'A\ud83d\ude00BC'],
  ['two-emoji', '\ud83d\ude00\ud83d\ude80A'],
]) {
  add('literal-' + name, 'literal', 1, [text + '\n\n'], name === 'emoji');
}

for (const [name, text] of [
  ['ascii-digits', '12345'],
  ['ascii-seeded', draw(ascii, 5)],
  ['bmp-seeded', draw(bmp, 5)],
  ['bmp-mixed', '\u6f22\u5b57\u03a9\u03b2A'],
  ['combining', 'e\u0301o\u0308A'],
  ['emoji', 'A\ud83d\ude00BC'],
  ['two-emoji', '\ud83d\ude00\ud83d\ude80A'],
]) {
  add('original-' + name, 'original', 1, [text + '\n\n'], name === 'bmp-seeded');
}

const label = parseVsd(inputs.original).pages[0].shapes.find(s => s.id === 3).text;
assert.ok(label.endsWith('\n'));
add('label-greek', 'original', 3, [label.replace('\u03a9', '\u03b2')], true);
add('label-ascii-seeded', 'original', 3, [draw(ascii, label.length - 1) + '\n']);
add('label-bmp-seeded', 'original', 3, [draw(bmp, label.length - 1) + '\n']);
add('literal-repeat', 'literal', 1, [
  'Jello\n\n',
  'A\ud83d\ude00BC\n\n',
  '\u03a9\u6f22\u5b57\u0430\u0431\n\n',
], true);
add('literal-edit-revert', 'literal', 1, ['Jello\n\n', 'Hello\n\n'], true);

const cases = [];
for (const p of plans) {
  const doc = parseVsd(inputs[p.fixture]);
  const shape = doc.pages[0].shapes.find(s => s.id === p.shapeId);
  const original = shape.text;
  let refusal = null;
  for (const text of p.words) {
    const old = Buffer.from(doc.serialize());
    const revision = doc.revision;
    const dirty = doc.dirty;
    try {
      shape.text = text;
    } catch (e) {
      assert.deepEqual(Buffer.from(doc.serialize()), old);
      assert.equal(doc.revision, revision);
      assert.equal(doc.dirty, dirty);
      refusal = { name: e.name, reason: e.reason ?? e.message };
      break;
    }
  }
  const bytes = Buffer.from(doc.serialize());
  const path = join(out, p.name + '.vsd');
  writeFileSync(path, bytes);
  cases.push({
    ...p,
    original,
    storedAfter: shape.text,
    nativeExpectedAfter: shape.text.slice(0, -1),
    revision: doc.revision,
    dirty: doc.dirty,
    output: path,
    sha256: sha(bytes),
    status: refusal ? 'REFUSED' : 'ADMITTED',
    refusal,
  });
}

const negatives = [];
for (const fixture of Object.keys(inputs)) {
  const doc = parseVsd(inputs[fixture]);
  const shape = doc.pages[0].shapes.find(s => s.id === 1);
  const source = inputs[fixture];
  for (const [name, text] of [
    ['short', 'Abcd\n\n'],
    ['long', 'Abcdef\n\n'],
    ['newline', 'Ab\nCD\n\n'],
    ['tab', 'Ab\tCD\n\n'],
    ['nul', 'Ab\0CD\n\n'],
    ['lone-high', 'A\ud83dBCD\n\n'],
    ['lone-low', 'A\ude00BCD\n\n'],
    ['terminal-control', 'Hello\r\n'],
  ]) {
    assert.throws(() => { shape.text = text; });
    assert.equal(doc.revision, 0);
    assert.equal(doc.dirty, false);
    assert.equal(shape.text, 'Hello\n\n');
    assert.deepEqual(Buffer.from(doc.serialize()), source);
    negatives.push({
      fixture, name, status: 'REFUSED-ATOMIC', revision: 0, dirty: false,
      serializedSha256: sha(source),
    });
  }
  const path = join(out, fixture + '-after-refusals.vsd');
  writeFileSync(path, Buffer.from(doc.serialize()));
}

for (const fixture of Object.keys(inputs)) {
  const doc = parseVsd(inputs[fixture]);
  const shape = doc.pages[0].shapes.find(s => s.id === 1);
  shape.text = 'World\n\n';
  const prior = Buffer.from(doc.serialize());
  const revision = doc.revision;
  const dirty = doc.dirty;
  for (const [name, text] of [
    ['short', 'Abcd\n\n'],
    ['long', 'Abcdef\n\n'],
    ['newline', 'Ab\nCD\n\n'],
    ['tab', 'Ab\tCD\n\n'],
    ['nul', 'Ab\0CD\n\n'],
    ['lone-high', 'A\ud83dBCD\n\n'],
    ['lone-low', 'A\ude00BCD\n\n'],
    ['terminal-control', 'Hello\r\n'],
  ]) {
    assert.throws(() => { shape.text = text; });
    assert.deepEqual(Buffer.from(doc.serialize()), prior);
    assert.equal(doc.revision, revision);
    assert.equal(doc.dirty, dirty);
    assert.equal(shape.text, 'World\n\n');
    negatives.push({
      fixture, name: 'dirty-' + name, status: 'REFUSED-ATOMIC', revision, dirty,
      serializedSha256: sha(prior),
    });
  }
}

const unsupported = parseVsd(inputs.original);
assert.throws(() => {
  unsupported.pages[0].shapes.find(s => s.id === 2).text = 'Hello\n\n';
});
assert.deepEqual(Buffer.from(unsupported.serialize()), inputs.original);
assert.equal(unsupported.revision, 0);
assert.equal(unsupported.dirty, false);
negatives.push({
  fixture: 'original', name: 'shape-without-explicit-text', status: 'REFUSED-ATOMIC',
  revision: 0, dirty: false, serializedSha256: sha(inputs.original),
});

const modules = readdirSync(dirname(modulePath))
  .filter(n => n.endsWith('.js'))
  .sort()
  .map(name => ({
    name,
    sha256: sha(readFileSync(join(dirname(modulePath), name))),
  }));
const report = {
  seed, seedHex: '0xc0deface', modulePath, modules, fixtures, cases, negatives,
  invalidCaseCount: negatives.length, nativeValidatedByThisRun: false,
};
writeFileSync(join(out, 'matrix-generation.json'), JSON.stringify(report, null, 2));
console.log(JSON.stringify({
  cases: cases.length,
  admitted: cases.filter(c => c.status === 'ADMITTED').length,
  refused: cases.filter(c => c.status === 'REFUSED').length,
  negativeAtomic: negatives.length,
}));
