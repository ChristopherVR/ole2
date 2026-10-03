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
 if(variant==='duplicate') word.set([0x3e,0x2a,1],blob+4);
 if(variant==='invalid') word[blob+3]=4;
 if(variant==='opaque') word[blob+4]=0x99;
 if(variant==='shared') word[slots.find(s=>word[s]!==word[slot])!]=word[slot]!;
 const table=cfb.tableBytes.slice();
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
 it('validates low-level inputs before inspection',()=>{for(const value of ['wavy',0,null])expect(()=>writeDocCharacterRunUnderline(new Uint8Array(),0,1,value as never)).toThrow('invalid-formatting');});
});
