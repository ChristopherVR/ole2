import { readFileSync } from 'node:fs';
import { describe, expect, it } from 'vitest';
import { DocDocument } from '../src/doc-document.js';
import { unwrapDocBytes } from '../src/ole-document-doc-cfb.js';
import { readDocFib } from '../src/ole-document-doc-fib.js';
import { parseBteTable } from '../src/ole-document-doc-fkp.js';
import { writeDocCharacterRunUnderline } from '../src/ole-document-doc-runs.js';
function prepared(variant = 'valid') {
 const input = new Uint8Array(readFileSync(new URL('./fixtures/doc/rich-runs.doc', import.meta.url)));
 const cfb = unwrapDocBytes(input)!, word = cfb.wordDocBytes.slice(), fib = readDocFib(word);
 const bte = parseBteTable(cfb.tableBytes, fib.plcfbteChpx), page = bte.pns[0]! * 512;
 const count = word[page + 511]!, rgb = page + (count + 1) * 4;
 const slots = Array.from({length: count}, (_, i) => rgb + i);
 const slot = slots.find(s => {const b = page + word[s]! * 2; return word.subarray(b+1,b+1+word[b]!).some((v,i,a)=>v===0x35 && a[i+1]===8);})!;
 const blob = page + word[slot]! * 2, length = word[blob]!;
 word.copyWithin(blob+4,blob+7,blob+1+length); word[blob]=length-3; word.set([0x3e,0x2a,1],blob+1);
 if(variant==='section-alias') new DataView(word.buffer).setUint16(blob-2,16,true);
 if(variant==='duplicate') word.set([0x3e,0x2a,1],blob+4);
 if(variant==='invalid') word[blob+3]=4;
 if(variant==='opaque') word[blob+4]=0x99;
 if(variant==='shared') word[slots.find(s=>word[s]!==word[slot])!]=word[slot]!;
 const table=cfb.tableBytes.slice();
 if(variant==='section-alias') new DataView(table.buffer).setInt32(fib.sed.fc+10,blob-2,true);
 if(variant==='reserved') new DataView(table.buffer).setUint32(fib.plcfbteChpx.fc+(bte.fcs.length)*4, bte.pns[0]! | 0x80000000,true);
 if(variant==='alias') new DataView(table.buffer).setUint32(fib.plcfbtePapx.fc+8,page/512 | 0x80000000,true);
 return cfb.rewrap(word,table);
}
describe('existing exclusive DOC underline',()=>{
 it('changes one byte, refreshes raw operand and restores every source byte',()=>{
  const input=prepared(), doc=new DocDocument(input), p=doc.paragraphs[1]!, run=p.runs[0]!;
  expect(run.directUnderline).toBe('single'); run.directUnderline='double';
  expect(run.directUnderline).toBe('double'); expect(run.sprms.find(p=>p.opcode===0x2a3e)!.operand).toEqual([3]);
  expect(doc.serialize().filter((v,i)=>v!==input[i]).length).toBe(1);
  expect(()=>{run.directUnderline='none';}).toThrow('stale-run');
  p.runs[0]!.directUnderline='none'; expect(p.runs[0]!.directUnderline).toBe('none');
  p.runs[0]!.directUnderline='single'; expect(doc.serialize()).toEqual(input);
 });
 it('keeps noops clean and rejects primitive-invalid or coercive inputs atomically',()=>{
  const input=prepared(),doc=new DocDocument(input),run=doc.paragraphs[1]!.runs[0]!;
  run.directUnderline='single';expect(doc.dirty).toBe(false);
  for(const value of [undefined,null,1,'wavy',new String('single'),{toString(){run.directBold=false;return 'double';}}]){
   expect(()=>{run.directUnderline=value as never;}).toThrow('invalid-formatting');expect(doc.serialize()).toEqual(input);expect(doc.dirty).toBe(false);
  }
 });
 it('refuses opaque, shared and unsupported existing operands',()=>{
  for(const variant of ['invalid','opaque','shared','duplicate']){const input=prepared(variant),doc=new DocDocument(input);expect(()=>{doc.paragraphs[1]!.runs[0]!.directUnderline='double';}).toThrow();expect(doc.serialize()).toEqual(input);}
 });
 it('ignores reserved CHPX PN bits without rewriting them and detects reserved-bit PAPX alias',()=>{
  const input=prepared('reserved'),doc=new DocDocument(input);expect(doc.paragraphs[1]!.runs[0]!.directUnderline).toBe('single');doc.paragraphs[1]!.runs[0]!.directUnderline='double';expect(doc.serialize().filter((v,i)=>v!==input[i]).length).toBe(1);
  const alias=prepared('alias'),bad=new DocDocument(alias);expect(()=>{bad.paragraphs[1]!.runs[0]!.directUnderline='double';}).toThrow('aliased-formatting');expect(bad.serialize()).toEqual(alias);
 });
 it('refuses missing operands and partial physical run views',()=>{
  const original=new Uint8Array(readFileSync(new URL('./fixtures/doc/rich-runs.doc',import.meta.url))); const doc=new DocDocument(original);
  expect(doc.paragraphs[1]!.runs[0]!.directUnderline).toBeUndefined();expect(()=>{doc.paragraphs[1]!.runs[0]!.directUnderline='single';}).toThrow('inherited-formatting');expect(doc.dirty).toBe(false);
  const input=prepared(),run=new DocDocument(input).paragraphs[1]!.runs[0]!;expect(()=>writeDocCharacterRunUnderline(input,run.cpStart+1,run.cpEnd,'double')).toThrow('unsupported-formatting');
 });
 it('refuses a valid section-exception pointer alias without changing any bytes',()=>{const input=prepared('section-alias'),doc=new DocDocument(input);expect(()=>{doc.paragraphs[1]!.runs[0]!.directUnderline='double';}).toThrow('aliased-formatting');expect(doc.serialize()).toEqual(input);expect(doc.dirty).toBe(false);});
 it('bounds aggregate CHPX operands before allocating unbounded stale rows',()=>{
  const cfb=unwrapDocBytes(prepared())!, fib=readDocFib(cfb.wordDocBytes), pages=3200, base=Math.ceil(cfb.wordDocBytes.length/512)*512;
  const word=new Uint8Array(base+pages*512);word.set(cfb.wordDocBytes);const table=new Uint8Array(cfb.tableBytes.length+pages*8+4);table.set(cfb.tableBytes);
  const wv=new DataView(word.buffer),tv=new DataView(table.buffer),fc=cfb.tableBytes.length;
  for(let i=0;i<=pages;i++)tv.setInt32(fc+i*4,2048+i*2,true);
  for(let i=0;i<pages;i++){const page=base+i*512;tv.setUint32(fc+(pages+1)*4+i*4,page/512,true);wv.setInt32(page,2048+i*2,true);wv.setInt32(page+4,2050+i*2,true);word[page+8]=5;word[page+10]=255;word[page+511]=1;for(let j=0;j<85;j++)word.set([0x35,8,0],page+11+j*3);}
  const descriptor=fib.fibRgFcLcbOffset+12*8;wv.setUint32(descriptor,fc,true);wv.setUint32(descriptor+4,pages*8+4,true);
  const input=cfb.rewrap(word,table), doc=new DocDocument(input);expect(()=>doc.paragraphs[1]!.runs).toThrow('unsupported-formatting');expect(()=>writeDocCharacterRunUnderline(input,23,32,'double')).toThrow('Formatting SPRM limit');expect(doc.dirty).toBe(false);expect(Buffer.compare(Buffer.from(doc.serialize()),Buffer.from(input))).toBe(0);
 });
 it('validates low-level inputs before inspection',()=>{for(const value of ['wavy',0,null])expect(()=>writeDocCharacterRunUnderline(new Uint8Array(),0,1,value as never)).toThrow('invalid-formatting');});
});
