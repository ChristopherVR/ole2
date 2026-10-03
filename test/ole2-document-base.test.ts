import { describe, expect, it } from 'vitest';
import { buildOle2 } from '../src/ole2-parser-write.js';
import { CfbDocument } from '../src/ole2-document.js';

describe('shared document base compatibility and ownership', () => {
 it('keeps getStream detachable, stable and current across multiple model edits', () => {
  const doc = new CfbDocument(new Uint8Array(buildOle2(new Map([['Example', new Uint8Array([1,2,3])]]))));
  const { getStream } = doc;
  expect(getStream).toBe(doc.getStream);
  expect(getStream('Example')).toEqual(new Uint8Array([1,2,3]));
  const handle = doc.stream(['Example']);
  handle.bytes = new Uint8Array([4,5,6,7]);
  expect(getStream('Example')).toEqual(new Uint8Array([4,5,6,7]));
  handle.bytes = new Uint8Array([8,9]);
  expect(getStream('Example')).toEqual(new Uint8Array([8,9]));
  expect(getStream('Missing')).toBeUndefined();
  getStream('Example')!.fill(0);
  expect(getStream('Example')).toEqual(new Uint8Array([8,9]));
  expect(doc.revision).toBe(2);
 });

 it('keeps input, serialization and nested directory metadata snapshots isolated', () => {
  const input = new Uint8Array(buildOle2(new Map([['Example', new Uint8Array([1,2,3])]])));
  const original = input.slice(), doc = new CfbDocument(input), { getStream } = doc;
  input.fill(0);
  doc.entries[0]!.clsid.fill(9);
  doc.entries[0]!.name = 'Replaced root snapshot';
  doc.serialize().fill(0);
  expect(doc.serialize()).toEqual(original);
  expect(getStream('Example')).toEqual(new Uint8Array([1,2,3]));
  expect(doc.dirty).toBe(false);
  expect(doc.revision).toBe(0);
 });
});
