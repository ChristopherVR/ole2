import { readFileSync } from 'node:fs';
import { isDeepStrictEqual } from 'node:util';
import { pathToFileURL } from 'node:url';

/** Compare independent consumer snapshots after explicitly declared field edits.
 * Paths use JSON Pointer syntax; every field outside the declared edits must match.
 * This is semantic evidence for the captured fields, never full native fidelity.
 */
export function compareNativeSnapshots(before, after, changes = []) {
  const expected = structuredClone(before);
  for (const { path, value } of changes) {
    if (typeof path !== 'string' || !path.startsWith('/') || path === '/') throw new Error('Invalid change path');
    const keys = path.slice(1).split('/').map(k => k.replace(/~1/g, '/').replace(/~0/g, '~'));
    if (keys.some(k => ['__proto__', 'prototype', 'constructor'].includes(k))) throw new Error('Unsafe change path');
    let parent = expected;
    for (const key of keys.slice(0, -1)) {
      if (parent === null || typeof parent !== 'object' || !Object.hasOwn(parent, key)) throw new Error(`Missing path ${path}`);
      parent = parent[key];
    }
    const key = keys.at(-1);
    if (parent === null || typeof parent !== 'object' || !Object.hasOwn(parent, key)) throw new Error(`Missing path ${path}`);
    parent[key] = value;
  }
  return { pass: isDeepStrictEqual(expected, after), evidence: 'native-consumer-semantic-snapshot', fullFidelity: false };
}

if (process.argv[1] && import.meta.url === pathToFileURL(process.argv[1]).href) {
  const [, , beforePath, afterPath, changesPath] = process.argv;
  if (!beforePath || !afterPath) throw new Error('Usage: node scripts/compare-native-snapshots.mjs BEFORE.json AFTER.json [CHANGES.json]');
  const read = path => JSON.parse(readFileSync(path, 'utf8').replace(/^\uFEFF/, ''));
  const result = compareNativeSnapshots(read(beforePath), read(afterPath), changesPath ? read(changesPath) : []);
  console.log(JSON.stringify(result));
  process.exitCode = result.pass ? 0 : 1;
}
