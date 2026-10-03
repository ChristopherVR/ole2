import{describe,it,expect}from'vitest';import{readFileSync}from'node:fs';import{VsdDocument}from'../src/vsd-document.js';import{readVsdDrawing}from'../src/vsd-reader.js';import{readCompoundFileStream}from'../src/ole2-stream-edit.js';
const input=()=>new Uint8Array(readFileSync(new URL('./fixtures/vsd/native-literal-transform.vsd',import.meta.url)));
describe('native text exact decoded-length safety',()=>{
 it.each(['Jello','World','Again','Hallo'])('edits %s without the decoded suffix that caused native rejection',word=>{
  const bytes=input(),before=bytes.slice(),old=readVsdDrawing(bytes),doc=new VsdDocument(bytes),shape=doc.pages[0]!.shapes[0]!;
  shape.text=word+'\n\n';const output=doc.serialize(),next=readVsdDrawing(output),record=old.pages[0]!.shapes[0]!.textRecord!,actual=next.pages[0]!.shapes[0]!.textRecord!.block.bytes.slice();
  expect(next.pages[0]!.shapes[0]!.text).toBe(word+'\n\n');expect(actual.length).toBe(record.block.bytes.length);
  actual.set(record.block.bytes.subarray(record.offset,record.offset+record.length),record.offset);expect(actual).toEqual(record.block.bytes);
  expect(output.length).toBe(bytes.length);expect(next.blocks.map(b=>[b.offset,b.length,b.format])).toEqual(old.blocks.map(b=>[b.offset,b.length,b.format]));
  expect(next.stream.subarray(0,record.block.offset)).toEqual(old.stream.subarray(0,record.block.offset));expect(next.stream.subarray(record.block.offset+record.block.length)).toEqual(old.stream.subarray(record.block.offset+record.block.length));
  for(const name of ['\u0005SummaryInformation','\u0005DocumentSummaryInformation'])expect(readCompoundFileStream(output,[name])).toEqual(readCompoundFileStream(bytes,[name]));
  expect(bytes).toEqual(before);expect(doc.revision).toBe(1);expect(doc.dirty).toBe(true);
 });
 it('retains handles across repeated edits and keeps rejected edits and no-ops atomic',()=>{
  const bytes=input(),doc=new VsdDocument(bytes),shape=doc.pages[0]!.shapes[0]!;shape.text=shape.text!;expect(doc.serialize()).toEqual(bytes);expect(doc.revision).toBe(0);
  for(const word of ['Jello','World','Again'])shape.text=word+'\n\n';expect(doc.revision).toBe(3);const before=doc.serialize();
  expect(()=>{shape.text='Longer\n\n';}).toThrow(/text-length-change/);expect(doc.serialize()).toEqual(before);expect(doc.revision).toBe(3);expect(shape.text).toBe('Again\n\n');
  shape.text=shape.text!;expect(doc.revision).toBe(3);expect(doc.serialize()).toEqual(before);
 });
});
