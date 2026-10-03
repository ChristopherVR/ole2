import { describe, expect, it } from 'vitest';
import { buildOle2 } from '../src/ole2-parser-write.js';
import { readCompoundFileStream } from '../src/ole2-stream-edit.js';
import { VsdDocument } from '../src/vsd-document.js';
const parseVsd = (input: Uint8Array) => new VsdDocument(input);
import { readVsdDrawing } from '../src/vsd-reader.js';
import { replaceVsdShapeText, replaceVsdShapeTransform } from '../src/vsd-writer.js';
import { readFileSync } from 'node:fs';
import { decodeVsdBlock, encodeVsdBlock } from '../src/vsd-compression.js';

import { makeVsdFixture } from './fixtures/vsd/generated.js';
const fixture = (compressed=true, mutate?: (stream:Uint8Array)=>void) => makeVsdFixture({compressed,mutate});
describe('binary VSD version 11 stored drawing model',()=>{
 it.each([false,true])('reads actual page/shape/text/transforms/line geometry compressed=%s', compressed=>{
  const doc=parseVsd(fixture(compressed)),page=doc.pages[0]!,shape=page.shapes[0]!;
  expect(doc.kind).toBe('vsd');expect(doc.version).toBe(11);expect(page.id).toBe(0);expect(page.width).toBe(8);expect(page.height).toBe(11);
  expect(shape.id).toBe(7);expect(shape.text).toBe('Hello\n');expect(shape.transform).toMatchObject({pinX:2,pinY:3,width:4,height:2});
  expect(shape.geometry).toEqual([{kind:'moveTo',id:0,x:0,y:0},{kind:'lineTo',id:1,x:4,y:2}]);expect(doc.dirty).toBe(false);
 });
 it.each([false,true])('edits text and transform through stale handles, preserves original unknown bytes compressed=%s', compressed=>{
  const input=fixture(compressed),original=readCompoundFileStream(input,['VisioDocument'])!,doc=parseVsd(input),shape=doc.pages[0]!.shapes[0]!,{getStream}=doc;
  shape.text='World\n';shape.transform={...shape.transform!,pinX:6,width:5};shape.text='Again\n';
  expect(shape.text).toBe('Again\n');expect(shape.transform?.pinX).toBe(6);expect(doc.revision).toBe(3);expect(doc.dirty).toBe(true);
  const output=doc.serialize(),reparsed=parseVsd(output),stream=readCompoundFileStream(output,['VisioDocument'])!;
  expect(reparsed.pages[0]!.shapes[0]!.text).toBe('Again\n');expect(reparsed.pages[0]!.shapes[0]!.transform?.width).toBe(5);
  expect(stream.length).toBe(original.length);
  const before=readVsdDrawing(input),after=readVsdDrawing(output),oldShape=before.pages[0]!.shapes[0]!,newShape=after.pages[0]!.shapes[0]!;
  const masked=newShape.textRecord!.block.bytes.slice();
  for(const record of [oldShape.textRecord!,oldShape.transformRecord!])masked.set(record.block.bytes.subarray(record.offset,record.offset+record.length),record.offset);
  expect(masked).toEqual(oldShape.textRecord!.block.bytes);
  expect(after.blocks.map(b=>[b.offset,b.length])).toEqual(before.blocks.map(b=>[b.offset,b.length]));
  expect(getStream('OpaqueUnknown')).toEqual(Uint8Array.of(9,4,8,3,5));
  expect(readCompoundFileStream(input,['VisioDocument'])).toEqual(original);
 });
 it('no-op stays clean; rejected edits leave bytes/model/revision unchanged',()=>{
  const doc=parseVsd(fixture()),shape=doc.pages[0]!.shapes[0]!,before=doc.serialize();shape.text='Hello\n';shape.transform={...shape.transform!};
  expect(doc.revision).toBe(0);expect(doc.dirty).toBe(false);
  expect(()=>{shape.text='Too long\n';}).toThrow('text-length-change');expect(()=>{shape.text='\ud800ello\n';}).toThrow();
  expect(()=>{shape.transform={...shape.transform!,width:NaN};}).toThrow('invalid-transform');expect(doc.serialize()).toEqual(before);expect(shape.text).toBe('Hello\n');
 });
 it.each([5,6,0,12])('refuses unsupported binary version %s',version=>expect(()=>parseVsd(fixture(false,s=>new DataView(s.buffer).setUint16(26,version,true)))).toThrow(`version-${version}`));
 it('rejects trailer outside document bounds and cyclic pointer graph',()=>{
  expect(()=>parseVsd(fixture(false,s=>new DataView(s.buffer).setUint32(44,0xffffffff,true)))).toThrow('pointer-bounds');
  expect(()=>parseVsd(fixture(false,s=>{const v=new DataView(s.buffer),root=v.getUint32(44,true);v.setUint32(root+28,root,true);}))).toThrow('pointer-cycle');
 });
 it('returned inputs, stream buffers and immutable handle snapshots cannot change owned state',()=>{
  const input=fixture(),doc=parseVsd(input),shape=doc.pages[0]!.shapes[0]!;input.fill(0);doc.getStream('VisioDocument')!.fill(0);doc.serialize().fill(0);
  expect(shape.text).toBe('Hello\n');expect(Object.isFrozen(shape.transform)).toBe(true);expect(Object.isFrozen(shape.geometry)).toBe(true);
  expect((shape as unknown as Record<string,unknown>).read).toBeUndefined();
 });
 it('exposes bounded graph details independently of typed handles',()=>{const d=readVsdDrawing(fixture());expect(d.blocks).toHaveLength(3);expect(d.writable).toBe(true);});
 it('retains a reproducible wholly authored fixture',()=>expect(new Uint8Array(readFileSync(new URL('./fixtures/vsd/owned-v11.vsd',import.meta.url)))).toEqual(fixture()));
 it('rejects moved structural text controls and ignores no writable runtime model identity',()=>{
  const doc=parseVsd(fixture());expect(()=>{doc.pages[0]!.shapes[0]!.text='Hell\no';}).toThrow('text-structure-change');
  expect(()=>{Object.assign(doc,{kind:'cfb'});}).toThrow();expect(()=>{Object.assign(doc,{pages:[]});}).toThrow();expect(doc.pages).toHaveLength(1);
 });
 it('retains a leading UTF16 BOM as a stored code unit and verifies exact text edits',()=>{
  const input=fixture(false,s=>{for(let i=0;i<s.length-12;i++)if(s[i]===72&&s[i+1]===0&&s[i+2]===101&&s[i+3]===0){s[i]=255;s[i+1]=254;break;}});
  expect(readVsdDrawing(input).pages[0]!.shapes[0]!.text).toBe('\ufeffello\n');
  expect(readVsdDrawing(replaceVsdShapeText(input,0,7,'\ufefforld\n')).pages[0]!.shapes[0]!.text).toBe('\ufefforld\n');
 });
 it('captures transform accessors exactly once and validates the captured value',()=>{
  let reads=0;const value={pinX:2,pinY:3,get width(){return ++reads===1?-1:4;},height:2,localPinX:2,localPinY:1,angle:0,flipX:false,flipY:false};
  expect(()=>replaceVsdShapeTransform(fixture(false),0,7,value)).toThrow('invalid-transform');expect(reads).toBe(1);
 });
 it('rejects a stale outer transform candidate while preserving a reentrant committed text edit',()=>{
  const doc=parseVsd(fixture()),shape=doc.pages[0]!.shapes[0]!,before=shape.transform!;
  const value={...before,get width(){shape.text='World\n';return 5;}};
  expect(()=>{shape.transform=value;}).toThrow('reentrant-edit');expect(shape.text).toBe('World\n');expect(shape.transform?.width).toBe(4);expect(doc.revision).toBe(1);
 });
 it('refuses a second pointer alias to a page tree without changing bytes',()=>{
  const original=fixture(false),stream=readCompoundFileStream(original,['VisioDocument'])!,root=readVsdDrawing(original).blocks[0]!;
  const alias=new Uint8Array(root.bytes.length+18);alias.set(root.bytes);alias.set(root.bytes.subarray(20,38),38);new DataView(alias.buffer).setInt32(12,2,true);
  const revised=new Uint8Array(stream.length+alias.length);revised.set(stream);revised.set(alias,stream.length);const v=new DataView(revised.buffer);v.setUint32(28,revised.length,true);v.setUint32(44,stream.length,true);v.setUint32(48,alias.length,true);
  const input=new Uint8Array(buildOle2(new Map([['VisioDocument',revised]]))),before=input.slice();
  expect(readVsdDrawing(input).writable).toBe(false);expect(()=>replaceVsdShapeText(input,0,7,'World\n')).toThrow('shared-or-overlapping-allocation');expect(input).toEqual(before);
 });
 it('refuses unsupported guide ownership instead of attributing guide text to a parent',()=>{
  const input=fixture(false,s=>{for(let cursor=54;cursor<s.length-19;cursor++){const v=new DataView(s.buffer);if(v.getUint32(cursor,true)===0x48 && v.getUint32(cursor+12,true)===54){v.setUint32(cursor,0x4d,true);break;}}});
  expect(()=>readVsdDrawing(input)).toThrow('unsupported-guide');
 });
 it('refuses guide-dependent transform contexts and unknown unit tags',()=>{
  const input=fixture(false,s=>{for(let cursor=54;cursor<s.length-19;cursor++){const v=new DataView(s.buffer);if(v.getUint32(cursor,true)===0x9b && v.getUint32(cursor+12,true)===65){s[cursor+19]=0xff;break;}}});
  const transform=parseVsd(fixture()).pages[0]!.shapes[0]!.transform!;
  expect(()=>replaceVsdShapeTransform(input,0,7,transform)).toThrow('unsupported-transform-units');
 });
 it('bounds input before extraction and rejects invalid caller decompression limits',()=>{
  expect(()=>readVsdDrawing(new Uint8Array(64*1024*1024+1))).toThrow('input-budget');
  for(const limit of [-1,NaN,Infinity,0.5,65*1024*1024])expect(()=>decodeVsdBlock(Uint8Array.of(1,0),limit)).toThrow('Invalid');
 });
 it('preserves exact unknown directory metadata across an in-place compressed text edit',()=>{
  const input=fixture(),v=new DataView(input.buffer),dir=(v.getUint32(48,true)+1)*512;let slot=-1;
  for(let offset=dir;offset<dir+512;offset+=128){const n=v.getUint16(offset+64,true);if(n>=2&&new TextDecoder('utf-16le').decode(input.subarray(offset,offset+n-2))==='OpaqueUnknown'){slot=offset;break;}}
  expect(slot).toBeGreaterThanOrEqual(0);for(let i=80;i<116;i++)input[slot+i]=(i*3)%256;const metadata=input.slice(slot,slot+128);
  const doc=parseVsd(input);doc.pages[0]!.shapes[0]!.text='World\n';expect(doc.serialize().subarray(slot,slot+128)).toEqual(metadata);
 });
});
describe('VSD compression bounds',()=>{
 it.each([0,1,7,8,9,4097])('roundtrips independently encoded literals length %s',length=>{const b=Uint8Array.from({length},(_,i)=>i%251);expect(decodeVsdBlock(encodeVsdBlock(b))).toEqual(b);});
 it('decodes dictionary references and refuses truncated tokens/output amplification',()=>{
  expect(decodeVsdBlock(Uint8Array.of(0,0,0))).toEqual(new Uint8Array(3));
  expect(()=>decodeVsdBlock(Uint8Array.of(0))).toThrow('flags');expect(()=>decodeVsdBlock(Uint8Array.of(0,1))).toThrow('reference');
  expect(()=>decodeVsdBlock(Uint8Array.of(0,0,15),4)).toThrow('limit');
 });
});
