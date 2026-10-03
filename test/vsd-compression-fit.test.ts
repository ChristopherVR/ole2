import {describe,it,expect} from 'vitest';
import {decodeVsdBlock} from '../src/vsd-compression.js';
import {encodeVsdBlockToSize} from '../src/vsd-compression-fit.js';
import {readVsdDrawing} from '../src/vsd-reader.js';
import {readFileSync} from 'node:fs';

function randomBytes(length:number):Uint8Array{let state=123456789;return Uint8Array.from({length},()=>{state^=state<<13;state^=state>>>17;state^=state<<5;return state&255;});}
describe('bounded existing-allocation VSD compression',()=>{
 it('fits the real native page while retaining its complete decoded prefix',()=>{
  const drawing=readVsdDrawing(new Uint8Array(readFileSync(new URL('./fixtures/vsd/native-visio16-v11.vsd',import.meta.url)))),record=drawing.pages[0]!.shapes[0]!.textRecord!,decoded=record.block.bytes.slice();
  const view=new DataView(decoded.buffer);for(let i=0;i<5;i++)view.setUint16(record.offset+8+i*2,'World'.charCodeAt(i),true);
  const encoded=encodeVsdBlockToSize(decoded,record.block.length)!;
  expect(encoded.bytes.length).toBe(record.block.length);expect(encoded.paddingBytes).toBeLessThanOrEqual(4096);
  const actual=decodeVsdBlock(encoded.bytes,decoded.length+encoded.paddingBytes);
  expect(actual.subarray(0,decoded.length)).toEqual(decoded);expect(actual.subarray(decoded.length).every(v=>v===0)).toBe(true);
 });
 it.each([1,7,8,9,4095,4096,4097,16384])('retains dictionary rollover/overlap and exact capacity length=%s',length=>{
  const input=Uint8Array.from({length},(_,i)=>i%11===0?17:0),capacity=Math.ceil(length/4)+32;
  const result=encodeVsdBlockToSize(input,capacity)!;expect(result).toBeDefined();expect(result.bytes.length).toBe(capacity);
  expect(decodeVsdBlock(result.bytes,input.length+result.paddingBytes).subarray(0,input.length)).toEqual(input);
 });
 it('refuses insufficient capacity and impractical padding without touching caller bytes',()=>{
  const bytes=randomBytes(1024),before=bytes.slice();expect(encodeVsdBlockToSize(bytes,10)).toBeUndefined();expect(bytes).toEqual(before);
  expect(encodeVsdBlockToSize(Uint8Array.of(1),100000)).toBeUndefined();
 });
 it('preflights input, output and integer bounds before allocating',()=>{
  for(const capacity of [-1,Infinity,NaN,0.5,Number.MAX_SAFE_INTEGER,8*1024*1024+1])expect(()=>encodeVsdBlockToSize(Uint8Array.of(1),capacity)).toThrow(/budget/);
  expect(()=>encodeVsdBlockToSize(new Uint8Array(8*1024*1024+1),1)).toThrow(/budget/);
  expect(()=>encodeVsdBlockToSize(null as unknown as Uint8Array,1)).toThrow(/budget/);
  expect(encodeVsdBlockToSize(new Uint8Array(),0)).toEqual({bytes:new Uint8Array(),paddingBytes:0});
 });
 it('handles a large incompressible workload within fixed allocation/search bounds',()=>{
  const bytes=randomBytes(256*1024),before=bytes.slice(),capacity=bytes.length+Math.ceil(bytes.length/8),result=encodeVsdBlockToSize(bytes,capacity)!;
  expect(result).toBeDefined();expect(result.bytes.length).toBe(capacity);expect(bytes).toEqual(before);
  expect(decodeVsdBlock(result.bytes,bytes.length+result.paddingBytes).subarray(0,bytes.length)).toEqual(bytes);
 });
 it('explicitly bounds adversarial hash-collision search work',()=>{
  const triples:number[][]=[];for(let a=0;a<256;a++)for(let b=0;b<256;b++){const c=(-((a*251+b)*251))&65535;if(c<256)triples.push([a,b,c]);}
  let state=4321;const bytes=new Uint8Array(512*1024);
  for(let i=0;i<bytes.length;i+=3){state^=state<<13;state^=state>>>17;state^=state<<5;bytes.set(triples[(state>>>0)%triples.length]!.slice(0,Math.min(3,bytes.length-i)),i);}
  const before=bytes.slice();expect(()=>encodeVsdBlockToSize(bytes,bytes.length+Math.ceil(bytes.length/8))).toThrow(/work-limit/);expect(bytes).toEqual(before);
 });
});
