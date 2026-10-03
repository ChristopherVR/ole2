import {readFileSync} from 'node:fs';
import {runInNewContext} from 'node:vm';
import {describe,it,expect} from 'vitest';
import {parseOle2,parseXls,parseDoc,parsePpt,CfbDocument,Ole2DocumentError,buildOle2} from '../src/index.js';

const fixture=(name:string)=>new Uint8Array(readFileSync(new URL(`./fixtures/${name}`,import.meta.url)));
describe('unified checked editable documents',()=>{
 it('discriminates models and checks wrong-format calls at runtime',()=>{
  const xls=fixture('xls/workbook-features.xls');
  const model=parseOle2(xls);expect(model.kind).toBe('xls');
  expect(parseOle2<'xls'>(xls,{expect:'xls'}).sheets.length).toBe(4);
  expect(()=>parsePpt(xls)).toThrow(Ole2DocumentError);
  expect(()=>parseDoc(xls)).toThrow(/Expected doc/);
  expect(parseOle2(xls,{expect:'cfb'}).kind).toBe('cfb');
 });
 it('preserves unsupported and ambiguous container inspection with diagnostics',()=>{
  const broken=buildOle2(new Map([['WordDocument',new Uint8Array([1,2,3])]]));
  const model=parseOle2(broken);expect(model.kind).toBe('cfb');
  if(model.kind!=='cfb')throw new Error('Expected container');
  expect(model.detectedFormat).toBe('doc');expect(model.diagnostics[0]?.code).toBe('unsupported-model');
  expect(model.getStream('WordDocument')).toEqual(new Uint8Array([1,2,3]));
  expect(()=>parseDoc(broken)).toThrow(/Cannot decode doc/);
  const ambiguous=buildOle2(new Map([['Workbook',new Uint8Array([1])],['WordDocument',new Uint8Array([2])]]));
  const result=parseOle2(ambiguous);expect(result).toBeInstanceOf(CfbDocument);
  if(result.kind!=='cfb')throw new Error('Expected container');
  expect(result.diagnostics[0]?.code).toBe('ambiguous-format');
  expect(()=>parseXls(ambiguous)).toThrow(/Multiple root format/);
 });
 it('owns NodeBuffer, byte-offset views and cross-realm ArrayBuffers',()=>{
  const original=buildOle2(new Map([['Example',new Uint8Array([1,2,3])]]));
  const input=Buffer.from(new Uint8Array(original));const model=parseOle2(input);input.fill(0);
  expect(model.getStream('Example')).toEqual(new Uint8Array([1,2,3]));
  const framed=new Uint8Array(original.byteLength+11);framed.set(new Uint8Array(original),7);
  expect(parseOle2(framed.subarray(7,7+original.byteLength)).getStream('Example')).toEqual(new Uint8Array([1,2,3]));
  const crossRealm=runInNewContext('new ArrayBuffer(size)',{size:original.byteLength}) as ArrayBuffer;
  new Uint8Array(crossRealm).set(new Uint8Array(original));
  expect(parseOle2(crossRealm).getStream('Example')).toEqual(new Uint8Array([1,2,3]));
 });
 it('detaches inspection aliases and mutates only through supported stream handles',()=>{
  const model=parseOle2(buildOle2(new Map([['Example',new Uint8Array([1,2,3])]])),{expect:'cfb'});
  const baseline=model.serialize();model.entries[0]!.clsid.fill(7);model.entries.length=0;
  model.getStream('Example')!.fill(0);model.serialize().fill(0);
  expect(model.serialize()).toEqual(baseline);expect(model.dirty).toBe(false);
  const handle=model.stream(['Example']);handle.bytes=new Uint8Array([1,2,3]);expect(model.dirty).toBe(false);
  handle.bytes=new Uint8Array([4,5,6,7]);expect(model.dirty).toBe(true);expect(model.revision).toBe(1);
  expect(handle.bytes).toEqual(new Uint8Array([4,5,6,7]));
  expect(()=>{(model as unknown as {kind:string}).kind='ppt';}).toThrow();
 });
 it('routes actual XLS model mutations through preservation and serialize',()=>{
  const model=parseXls(fixture('xls/workbook-features.xls'));
  const cell=model.sheets[0]!.cell(1,0);cell.value='Model Unicode \u03a9 \u65e5\u672c';
  expect(cell.value).toBe('Model Unicode \u03a9 \u65e5\u672c');
  expect(parseXls(model.serialize()).sheets[0]!.cell(1,0).value).toBe(cell.value);
  expect(model.recalculationRequired).toBe(true);
 });
});
