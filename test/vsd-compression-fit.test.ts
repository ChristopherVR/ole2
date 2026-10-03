import {describe,it,expect} from 'vitest';
import {decodeVsdBlock,encodeVsdBlock} from '../src/vsd-compression.js';
import {encodeVsdBlockToSize} from '../src/vsd-compression-fit.js';
import {readVsdDrawing} from '../src/vsd-reader.js';
import {readFileSync} from 'node:fs';

function randomBytes(length:number):Uint8Array{let state=123456789;return Uint8Array.from({length},()=>{state^=state<<13;state^=state>>>17;state^=state<<5;return state&255;});}
describe('bounded existing-allocation VSD compression',()=>{
 it('fits the real native page while retaining its complete decoded prefix',()=>{
  const drawing=readVsdDrawing(new Uint8Array(readFileSync(new URL('./fixtures/vsd/native-visio16-v11.vsd',import.meta.url)))),record=drawing.pages[0]!.shapes[0]!.textRecord!,decoded=record.block.bytes.slice();
  const view=new DataView(decoded.buffer);for(let i=0;i<5;i++)view.setUint16(record.offset+8+i*2,'World'.charCodeAt(i),true);
  const encoded=encodeVsdBlockToSize(decoded,record.block.length)!;
  expect(encoded.bytes.length).toBe(record.block.length);expect(encoded.paddingBytes).toBe(0);
  const actual=decodeVsdBlock(encoded.bytes,decoded.length);
  expect(actual).toEqual(decoded);
 });
 it.each([1,7,8,9,4095,4096,4097,16384])('retains dictionary rollover/overlap and exact capacity length=%s',length=>{
  const input=Uint8Array.from({length},(_,i)=>i%11===0?17:0),capacity=length<10?encodeVsdBlock(input).length:Math.ceil(length/4)+32;
  const result=encodeVsdBlockToSize(input,capacity)!;expect(result).toBeDefined();expect(result.bytes.length).toBe(capacity);
  expect(decodeVsdBlock(result.bytes,input.length+result.paddingBytes).subarray(0,input.length)).toEqual(input);
 });
 it('rewrites the aligned tail after more than 4096 tokens without changing decoded bytes',()=>{
  const seed=randomBytes(2048),input=new Uint8Array(65536);for(let i=0;i<32;i++)input.set(seed,i*2048);
  const before=input.slice(),fit=encodeVsdBlockToSize(input,12000)!;
  expect(fit.bytes.length).toBe(12000);expect(fit.paddingBytes).toBe(0);expect(decodeVsdBlock(fit.bytes,input.length)).toEqual(input);expect(input).toEqual(before);
  let cursor=0,count=0;while(cursor<fit.bytes.length){const flags=fit.bytes[cursor++]!;for(let bit=0;bit<8&&cursor<fit.bytes.length;bit++){cursor+=(flags>>bit)&1?1:2;count++;}}
  expect(count).toBeGreaterThan(4096);
  expect(encodeVsdBlockToSize(input,14000)).toBeUndefined();
 });
 it('refuses insufficient capacity and impossible exact storage without touching caller bytes',()=>{
  const bytes=randomBytes(1024),before=bytes.slice();expect(encodeVsdBlockToSize(bytes,10)).toBeUndefined();expect(bytes).toEqual(before);
  expect(encodeVsdBlockToSize(Uint8Array.of(1),100000)).toBeUndefined();
 });
 it('preflights input, output and integer bounds before allocating',()=>{
  for(const capacity of [-1,Infinity,NaN,0.5,Number.MAX_SAFE_INTEGER,8*1024*1024+1])expect(()=>encodeVsdBlockToSize(Uint8Array.of(1),capacity)).toThrow(/budget/);
  expect(()=>encodeVsdBlockToSize(new Uint8Array(8*1024*1024+1),1)).toThrow(/budget/);
  expect(()=>encodeVsdBlockToSize(null as unknown as Uint8Array,1)).toThrow(/budget/);
  expect(encodeVsdBlockToSize(new Uint8Array(),0)).toEqual({bytes:new Uint8Array(),paddingBytes:0});
 });
 it('refuses a large incompressible allocation when bounded tail token variation cannot fit it',()=>{
  const bytes=randomBytes(256*1024),before=bytes.slice(),capacity=bytes.length+Math.ceil(bytes.length/8),result=encodeVsdBlockToSize(bytes,capacity)!;
  expect(result).toBeUndefined();expect(bytes).toEqual(before);
 });
 it('explicitly bounds adversarial hash-collision search work',()=>{
  const triples:number[][]=[];for(let a=0;a<256;a++)for(let b=0;b<256;b++){const c=(-((a*251+b)*251))&65535;if(c<256)triples.push([a,b,c]);}
  let state=4321;const bytes=new Uint8Array(512*1024);
  for(let i=0;i<bytes.length;i+=3){state^=state<<13;state^=state>>>17;state^=state<<5;bytes.set(triples[(state>>>0)%triples.length]!.slice(0,Math.min(3,bytes.length-i)),i);}
  const before=bytes.slice();expect(()=>encodeVsdBlockToSize(bytes,bytes.length+Math.ceil(bytes.length/8))).toThrow(/work-limit/);expect(bytes).toEqual(before);
 });
});
