import { readFileSync } from 'node:fs';
import { describe, it, expect } from 'vitest';
import { parseOle2, parseVsd, parseDoc, VsdDocument, Ole2DocumentBase, CfbDocument, Ole2DocumentError, buildOle2, readCompoundFileStream } from '../src/index.js';
import { makeVsdFixture } from './fixtures/vsd/generated.js';

describe('checked binary Visio dispatch', () => {
 it('dispatches actual owned drawing bytes through the primary model API', () => {
  const bytes = new Uint8Array(readFileSync(new URL('./fixtures/vsd/owned-v11.vsd', import.meta.url)));
  const document = parseOle2(bytes);
  expect(document).toBeInstanceOf(VsdDocument); expect(document).toBeInstanceOf(Ole2DocumentBase);
  if(document.kind !== 'vsd') throw new Error('Expected VSD model');
  const shape = document.pages[0]!.shapes[0]!;
  shape.text = 'World\n'; shape.transform = {...shape.transform!, pinX: 6, width: 5};
  expect(parseVsd(document.serialize()).pages[0]!.shapes[0]!.text).toBe('World\n');
  expect(document.getStream('OpaqueUnknown')).toEqual(Uint8Array.of(9,4,8,3,5));
  expect(parseOle2(bytes, {expect:'vsd'}).version).toBe(11);
  expect(parseOle2(bytes, {expect:'cfb'})).toBeInstanceOf(CfbDocument);
  expect(() => parseDoc(bytes)).toThrow(/Expected doc/);
 });
 it.each([5,6,12])('retains unsupported version %s as diagnostic CFB inspection', version => {
  const bytes = makeVsdFixture({mutate: stream => new DataView(stream.buffer).setUint16(26,version,true)});
  const document = parseOle2(bytes);
  expect(document.kind).toBe('cfb');
  if(document.kind !== 'cfb') throw new Error('Expected diagnostic container');
  expect(document.detectedFormat).toBe('vsd'); expect(document.diagnostics[0]?.code).toBe('unsupported-model');
  expect(document.serialize()).toEqual(bytes);
  expect(() => parseVsd(bytes)).toThrow(Ole2DocumentError);
 });
 it('rejects mixed root identities and unrelated compound files', () => {
  const source = makeVsdFixture();
  const vsd = readCompoundFileStream(source,['VisioDocument'])!;
  const mixed = buildOle2(new Map([['VisioDocument',vsd],['WordDocument',Uint8Array.of(1)]]));
  const document = parseOle2(mixed);
  expect(document.kind).toBe('cfb');
  if(document.kind !== 'cfb') throw new Error('Expected diagnostic container');
  expect(document.diagnostics[0]?.code).toBe('ambiguous-format');
  expect(() => parseVsd(mixed)).toThrow(/Multiple root format/);
  const unrelated = buildOle2(new Map([['Data',Uint8Array.of(1)]]));
  expect(() => parseVsd(unrelated)).toThrow(/Expected vsd/);
 });
});
