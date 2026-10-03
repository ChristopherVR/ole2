import {describe,it,expect} from 'vitest';
import {VsdDocument} from '../src/vsd-document.js';
import {readVsdDrawing} from '../src/vsd-reader.js';
import {readCompoundFileStream} from '../src/ole2-stream-edit.js';
import {makeVsdFixture} from './fixtures/vsd/generated.js';
import {readFileSync} from 'node:fs';
import {replaceVsdShapeText} from '../src/vsd-writer.js';
import {UnsupportedOle2EditError} from '../src/ole2-document-base.js';
import {encodeVsdBlockToSize} from '../src/vsd-compression-fit.js';
import {resizeCompoundFileStream} from '../src/ole2-stream-resize.js';
import {parseOle2} from '../src/ole2-parser-read.js';
import {encodeVsdBlock} from '../src/vsd-compression.js';
import {buildOle2} from '../src/ole2-parser-write.js';

describe('VSD writers retain original block locations',()=>{
 it.each([false,true])('changes only the selected leaf and retains all pointers compressed=%s',compressed=>{
  const input=makeVsdFixture({compressed}),before=readVsdDrawing(input),doc=new VsdDocument(input);
  const record=before.pages[0]!.shapes[0]!.textRecord!;
  doc.pages[0]!.shapes[0]!.text='World\n';
  const after=readVsdDrawing(doc.serialize());
  expect(after.stream.length).toBe(before.stream.length);
  expect(after.blocks.map(b=>[b.type,b.offset,b.length,b.format])).toEqual(before.blocks.map(b=>[b.type,b.offset,b.length,b.format]));
  expect(after.stream.subarray(0,record.block.offset)).toEqual(before.stream.subarray(0,record.block.offset));
  expect(after.stream.subarray(record.block.offset+record.block.length)).toEqual(before.stream.subarray(record.block.offset+record.block.length));
  expect(after.pages[0]!.shapes[0]!.text).toBe('World\n');
  expect(readCompoundFileStream(doc.serialize(),['OpaqueUnknown'])).toEqual(readCompoundFileStream(input,['OpaqueUnknown']));
 });
 it('edits a real native-compressed page without relocating its allocation or pointer graph',()=>{
  const input=new Uint8Array(readFileSync(new URL('./fixtures/vsd/native-visio16-v11.vsd',import.meta.url))),before=input.slice(),doc=new VsdDocument(input),shape=doc.pages[0]!.shapes.find(s=>s.id===1)!;
  expect(shape.text).toBe('Hello\n\n');
  expect(replaceVsdShapeText(input,0,1,'Hello\n\n')).toEqual(before);shape.text=shape.text!;
  expect(doc.serialize()).toEqual(before);expect(doc.revision).toBe(0);expect(doc.dirty).toBe(false);
  shape.text='World\n\n';expect(doc.serialize()).toEqual(replaceVsdShapeText(input,0,1,'World\n\n'));
  const oldDrawing=readVsdDrawing(input),newDrawing=readVsdDrawing(doc.serialize());
  expect(newDrawing.blocks.map(b=>[b.type,b.offset,b.length,b.format])).toEqual(oldDrawing.blocks.map(b=>[b.type,b.offset,b.length,b.format]));
  const block=oldDrawing.pages[0]!.shapes[0]!.textRecord!.block,oldStream=oldDrawing.stream,newStream=newDrawing.stream;
  expect(newStream.subarray(0,block.offset)).toEqual(oldStream.subarray(0,block.offset));expect(newStream.subarray(block.offset+block.length)).toEqual(oldStream.subarray(block.offset+block.length));
  const oldLeaf=block.bytes,actualLeaf=newDrawing.pages[0]!.shapes[0]!.textRecord!.block.bytes.slice(0,oldLeaf.length),textRecord=oldDrawing.pages[0]!.shapes[0]!.textRecord!;
  actualLeaf.set(oldLeaf.subarray(textRecord.offset,textRecord.offset+textRecord.length),textRecord.offset);expect(actualLeaf).toEqual(oldLeaf);
  // Map only the existing regular-stream leaf allocation to physical bytes.
  // Every header/directory/property/unused byte outside it must remain exact.
  const view=new DataView(input.buffer),fatSector=view.getUint32(76,true),fat=(fatSector+1)*512;
  const entry=parseOle2(input.buffer).entries.find(e=>e.name==='VisioDocument')!,allowed=new Set<number>();let sector=entry.startSector,logical=0;
  while(sector<0xfffffffa){for(let i=0;i<512;i++)if(logical+i>=block.offset&&logical+i<block.offset+block.length)allowed.add((sector+1)*512+i);logical+=512;sector=view.getUint32(fat+sector*4,true);}
  const output=doc.serialize();for(let i=0;i<input.length;i++)if(!allowed.has(i))expect(output[i]).toBe(input[i]);
  expect(doc.serialize().length).toBe(input.length);expect(input).toEqual(before);expect(shape.text).toBe('World\n\n');expect(doc.revision).toBe(1);
  const committed=doc.serialize();shape.text=shape.text!;expect(doc.serialize()).toEqual(committed);expect(doc.revision).toBe(1);
  expect(()=>{shape.text='World\r\n';}).toThrow(UnsupportedOle2EditError);expect(doc.serialize()).toEqual(committed);
  shape.text='Again\n\n';expect(shape.text).toBe('Again\n\n');expect(doc.revision).toBe(2);expect(doc.serialize().length).toBe(input.length);
 });
 it('refuses an authored capacity-constrained variant atomically rather than relocating',()=>{
  const input=new Uint8Array(readFileSync(new URL('./fixtures/vsd/native-visio16-v11.vsd',import.meta.url))),drawing=readVsdDrawing(input),block=drawing.pages[0]!.shapes[0]!.textRecord!.block;
  const fit=encodeVsdBlockToSize(block.bytes,1210)!;expect(fit.paddingBytes).toBe(0);
  const stream=drawing.stream.slice();stream.set(fit.bytes,block.offset);
  // Mechanical bounds fixture, not an additional native-produced document:
  // constrain the parsed leaf range while retaining all its decoded bytes.
  new DataView(stream.buffer).setUint32(block.parent!.offset+block.pointerOffset+12,fit.bytes.length,true);
  const resized=resizeCompoundFileStream(input,['VisioDocument'],stream);if(!resized.ok)throw new Error(resized.reason);
  const doc=new VsdDocument(resized.bytes),shape=doc.pages[0]!.shapes[0]!,before=doc.serialize();
  expect(shape.text).toBe('Hello\n\n');expect(()=>{shape.text='World\n\n';}).toThrow(/unsafe-block-relocation/);
  expect(doc.serialize()).toEqual(before);expect(doc.dirty).toBe(false);expect(doc.revision).toBe(0);expect(shape.text).toBe('Hello\n\n');
 });
 it('refuses an oversized decoded compressed leaf before allocating writer buffers',()=>{
  const drawing=readVsdDrawing(makeVsdFixture({compressed:false})),block=drawing.pages[0]!.shapes[0]!.textRecord!.block;
  // Authored mechanical bounds fixture: valid chunks plus a large zero suffix,
  // encoded with independent direct LZSS zero references. No native claim.
  const prefix=new Uint8Array(Math.ceil((block.bytes.length+8)/8)*8);prefix.set(block.bytes);
  const literals=encodeVsdBlock(prefix),groups=Math.ceil((8*1024*1024+1-prefix.length)/144),stored=new Uint8Array(literals.length+groups*17);stored.set(literals);
  let cursor=literals.length,position=prefix.length;
  for(let g=0;g<groups;g++){stored[cursor++]=0;for(let token=0;token<8;token++){const address=(position-26)&4095;stored[cursor++]=address&255;stored[cursor++]=((address>>4)&240)|15;position+=18;}}
  const stream=new Uint8Array(drawing.stream.length+stored.length);stream.set(drawing.stream);stream.set(stored,drawing.stream.length);
  const view=new DataView(stream.buffer),pointer=block.parent!.offset+block.pointerOffset;view.setUint32(pointer+8,drawing.stream.length,true);view.setUint32(pointer+12,stored.length,true);view.setUint16(pointer+16,block.format|2,true);view.setUint32(28,stream.length,true);
  const input=new Uint8Array(buildOle2(new Map([['VisioDocument',stream]]))),doc=new VsdDocument(input),before=doc.serialize();
  expect(doc.pages[0]!.shapes[0]!.text).toBe('Hello\n');
  expect(()=>{doc.pages[0]!.shapes[0]!.text='World\n';}).toThrow(/compression-input-budget/);
  expect(doc.serialize()).toEqual(before);expect(doc.revision).toBe(0);expect(doc.dirty).toBe(false);
 });
});
