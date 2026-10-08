import { describe, expect, it } from 'vitest';
import { readFileSync } from 'node:fs';
import { buildOle2 } from '../src/ole2-parser-write.js';
import { encodeVsdBlock } from '../src/vsd-compression.js';
import { readVsdDrawing } from '../src/vsd-reader.js';
import { VsdDocument } from '../src/vsd-document.js';

// Authored record bytes based on libvisio VSDParser readGeometry/readArcTo.
// The fixture does not establish native acceptance or evaluate stored formulas.
function chunk(type:number,id:number,body:Uint8Array,level=3):Uint8Array {
 const bytes=new Uint8Array(19+body.length),v=new DataView(bytes.buffer);
 v.setUint32(0,type,true);v.setUint32(4,id,true);v.setUint32(12,body.length,true);v.setUint16(16,level,true);bytes[18]=0x50;bytes.set(body,19);return bytes;
}
function coordinates(values:number[]):Uint8Array {
 const bytes=new Uint8Array(values.length*9),v=new DataView(bytes.buffer);values.forEach((value,i)=>v.setFloat64(i*9+1,value,true));return bytes;
}
function fixture(compressed:boolean,rows:Uint8Array[]):Uint8Array {
 const shape=new Uint8Array(54),v=new DataView(shape.buffer);for(const offset of [10,18,26])v.setUint32(offset,0xffffffff,true);
 const chunks=[chunk(0x48,7,shape,1),...rows],page=new Uint8Array(chunks.reduce((length,c)=>length+c.length,0));let position=0;
 for(const c of chunks){page.set(c,position);position+=c.length;}
 const parts:Uint8Array[]=[];let size=54;
 function add(body:Uint8Array){const stored=compressed?encodeVsdBlock(body):body,range={offset:size,length:stored.length};parts.push(stored);size+=stored.length;return range;}
 function table(type:number,range:{offset:number;length:number},format:number){
  const shift=compressed?4:0,bytes=new Uint8Array(shift+38),v=new DataView(bytes.buffer);v.setUint32(shift,12,true);v.setInt32(shift+12,1,true);
  v.setUint32(shift+20,type,true);v.setUint32(shift+28,range.offset,true);v.setUint32(shift+32,range.length,true);v.setUint16(shift+36,format+(compressed?2:0),true);return bytes;
 }
 const pageRange=add(page),pagesRange=add(table(0x15,pageRange,0xd1)),trailerRange=add(table(0x27,pagesRange,0x50));
 const stream=new Uint8Array(size),header=new DataView(stream.buffer);stream.set(new TextEncoder().encode('Visio (TM) Drawing\r\n'));header.setUint16(26,11,true);header.setUint32(28,size,true);
 header.setUint32(36,0x14,true);header.setUint32(44,trailerRange.offset,true);header.setUint32(48,trailerRange.length,true);header.setUint16(52,0x50+(compressed?2:0),true);
 position=54;for(const part of parts){stream.set(part,position);position+=part.length;}
 return new Uint8Array(buildOle2(new Map([['VisioDocument',stream]])));
}
describe('stored VSD geometry sections',()=>{
 it.each([false,true])('retains section IDs, visibility flags and arc operands compressed=%s',compressed=>{
  const input=fixture(compressed,[chunk(0x89,3,Uint8Array.of(1)),chunk(0x8a,0,coordinates([0,0])),chunk(0x8c,1,coordinates([4,2,-0.5])),chunk(0x89,9,Uint8Array.of(6)),chunk(0x8b,8,coordinates([7,8]))]);
  const drawing=readVsdDrawing(input),shape=drawing.pages[0]!.shapes[0]!;
  expect(shape.geometrySections).toEqual([
   {id:3,noFill:true,noLine:false,noShow:false,unsupported:false,rows:[{kind:'moveTo',id:0,x:0,y:0},{kind:'arcTo',id:1,x:4,y:2,bow:-0.5}]},
   {id:9,noFill:false,noLine:true,noShow:true,unsupported:false,rows:[{kind:'lineTo',id:8,x:7,y:8}]}]);
  expect(shape.geometry).toEqual([{kind:'moveTo',id:0,x:0,y:0},{kind:'lineTo',id:8,x:7,y:8}]);
  expect(shape.unsupportedGeometry).toBe(true);expect(new VsdDocument(input).serialize()).toEqual(input);
 });
 it.each([false,true])('reads six stored ellipse operands without promoting flat geometry compressed=%s',compressed=>{
  const input=fixture(compressed,[chunk(0x89,0,Uint8Array.of(0)),chunk(0x8f,4,coordinates([4,3,1,2,5,7]))]);
  const doc=new VsdDocument(input),shape=doc.pages[0]!.shapes[0]!,section=shape.geometrySections[0]!;
  expect(section.rows).toEqual([{kind:'ellipse',id:4,centerX:4,centerY:3,leftX:1,leftY:2,topX:5,topY:7}]);
  expect(section.unsupported).toBe(false);expect(shape.unsupportedGeometry).toBe(true);expect(shape.geometry).toEqual([]);
  expect(Object.isFrozen(section.rows[0])).toBe(true);expect(Reflect.set(section.rows[0]!, 'centerX',99)).toBe(false);
  expect(shape.geometrySections[0]!.rows[0]).toEqual(section.rows[0]);expect(doc.serialize()).toEqual(input);expect(doc.revision).toBe(0);
 });
 it.each([false,true])('rejects truncated or nonfinite ellipse operands compressed=%s',compressed=>{
  expect(()=>readVsdDrawing(fixture(compressed,[chunk(0x89,0,Uint8Array.of(0)),chunk(0x8f,0,new Uint8Array(53))]))).toThrow('geometry-record');
  for(let index=0;index<6;index++)for(const value of [NaN,Infinity,-Infinity]){
   const values=[1,2,3,4,5,6];values[index]=value;
   expect(()=>readVsdDrawing(fixture(compressed,[chunk(0x89,0,Uint8Array.of(0)),chunk(0x8f,0,coordinates(values))]))).toThrow('non-finite-geometry');
  }
 });
 it('does not attach a new list without a header to the preceding section',()=>{
  const shape=readVsdDrawing(fixture(false,[chunk(0x89,0,Uint8Array.of(0)),chunk(0x8a,0,coordinates([1,2])),chunk(0x6c,1,new Uint8Array(),2),chunk(0x8b,1,coordinates([3,4]))])).pages[0]!.shapes[0]!;
  expect(shape.geometrySections[0]!.rows).toEqual([{kind:'moveTo',id:0,x:1,y:2}]);expect(shape.geometry).toHaveLength(2);
 });
 it('keeps nested and sibling shape rows in their own section contexts',()=>{
  const childHeader=new Uint8Array(54),headerView=new DataView(childHeader.buffer);headerView.setUint32(10,7,true);headerView.setUint32(18,0xffffffff,true);headerView.setUint32(26,0xffffffff,true);
  const rows=[chunk(0x89,10,Uint8Array.of(0)),chunk(0x8a,0,coordinates([1,2])),chunk(0x48,8,childHeader,2),chunk(0x8b,0,coordinates([3,4])),chunk(0x89,11,Uint8Array.of(0)),chunk(0x8a,1,coordinates([5,6])),chunk(0x48,9,childHeader,2),chunk(0x8b,2,coordinates([7,8])),chunk(0x89,12,Uint8Array.of(0)),chunk(0x8a,3,coordinates([9,10]))];
  const input=fixture(false,rows),doc=new VsdDocument(input),shapes=doc.pages[0]!.shapes;
  expect(shapes.map(s=>s.geometrySections.map(section=>({id:section.id,rows:section.rows})))).toEqual([
   [{id:10,rows:[{kind:'moveTo',id:0,x:1,y:2}]}],
   [{id:11,rows:[{kind:'moveTo',id:1,x:5,y:6}]}],
   [{id:12,rows:[{kind:'moveTo',id:3,x:9,y:10}]}]]);
  const first=shapes[0]!.geometrySections,second=shapes[0]!.geometrySections;
  expect(first).not.toBe(second);expect(first[0]).not.toBe(second[0]);expect(first[0]!.rows).not.toBe(second[0]!.rows);expect(first[0]!.rows[0]).not.toBe(second[0]!.rows[0]);
  expect(Reflect.set(first[0]!.rows[0]!, 'x',99)).toBe(false);expect(shapes[0]!.geometrySections[0]!.rows[0]).toMatchObject({x:1});
  expect(doc.serialize()).toEqual(input);
 });
 it('marks unknown section flags and undecoded rows without inventing geometry',()=>{
  const shape=readVsdDrawing(fixture(false,[chunk(0x89,0,Uint8Array.of(8)),chunk(0x8d,1,coordinates([1,2,3,4,5,6]))])).pages[0]!.shapes[0]!;
  expect(shape.geometrySections[0]).toMatchObject({unsupported:true,rows:[]});expect(shape.unsupportedGeometry).toBe(true);
 });
 it.each([false,true])('rejects truncated section/arc and nonfinite operands compressed=%s',compressed=>{
  for(const row of [chunk(0x89,0,new Uint8Array()),chunk(0x8c,0,new Uint8Array(26)),chunk(0x8c,0,coordinates([1,2,Infinity]))])
   expect(()=>readVsdDrawing(fixture(compressed,[chunk(0x89,0,Uint8Array.of(0)),row]))).toThrow(/geometry-section-record|geometry-record|non-finite-geometry/);
 });
 it('retains native hidden-section flags and byte-identical no-op serialization',()=>{
  const bytes=new Uint8Array(readFileSync(new URL('./fixtures/vsd/native-hierarchy.vsd',import.meta.url))),drawing=readVsdDrawing(bytes);
  expect(drawing.pages[1]!.shapes[0]!.geometrySections[0]).toMatchObject({id:0,noShow:true,rows:[{kind:'moveTo',id:1,x:0,y:0},{kind:'lineTo',id:2,x:4,y:0},{kind:'lineTo',id:3,x:4,y:2},{kind:'lineTo',id:4,x:0,y:2},{kind:'lineTo',id:5,x:0,y:0}]});
  const doc=new VsdDocument(bytes),sections=doc.pages[1]!.shapes[0]!.geometrySections;
  expect(Object.isFrozen(sections)).toBe(true);expect(Object.isFrozen(sections[0])).toBe(true);expect(Object.isFrozen(sections[0]!.rows)).toBe(true);expect(Object.isFrozen(sections[0]!.rows[0])).toBe(true);
  expect(doc.serialize()).toEqual(bytes);
 });
});
