import {describe,it,expect} from 'vitest';
import {VsdDocument} from '../src/vsd-document.js';
import {readVsdDrawing} from '../src/vsd-reader.js';
import {readCompoundFileStream} from '../src/ole2-stream-edit.js';
import {makeVsdFixture} from './fixtures/vsd/generated.js';
import {readFileSync} from 'node:fs';
import {replaceVsdShapeText} from '../src/vsd-writer.js';
import {UnsupportedOle2EditError} from '../src/ole2-document-base.js';

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
 it('refuses relocation of a real native-compressed page without changing owned state',()=>{
  const input=new Uint8Array(readFileSync(new URL('./fixtures/vsd/native-visio16-v11.vsd',import.meta.url))),before=input.slice(),doc=new VsdDocument(input),shape=doc.pages[0]!.shapes.find(s=>s.id===1)!;
  expect(shape.text).toBe('Hello\n\n');
  expect(()=>replaceVsdShapeText(input,0,1,'World\n\n')).toThrow(/unsafe-block-relocation/);
  expect(()=>{shape.text='World\n\n';}).toThrow(UnsupportedOle2EditError);
  expect(()=>{shape.text='World\n\n';}).toThrow(/unsafe-block-relocation/);
  expect(input).toEqual(before);expect(doc.serialize()).toEqual(before);expect(doc.revision).toBe(0);expect(doc.dirty).toBe(false);expect(shape.text).toBe('Hello\n\n');
  shape.text=shape.text!;
  expect(replaceVsdShapeText(input,0,1,shape.text!)).toEqual(before);
  expect(doc.serialize()).toEqual(before);expect(doc.revision).toBe(0);expect(doc.dirty).toBe(false);
 });
});
