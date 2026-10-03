import {readFileSync} from 'node:fs';
import {describe,it,expect} from 'vitest';
import {DocDocument} from '../src/doc-document.js';
import {UnsupportedOle2EditError} from '../src/ole2-document-base.js';
import {unwrapDocBytes} from '../src/ole-document-doc-cfb.js';
import {readDocFib} from '../src/ole-document-doc-fib.js';
import {parseBteTable} from '../src/ole-document-doc-fkp.js';
import {readDocCharacterRuns,writeDocCharacterRunFontSize} from '../src/ole-document-doc-runs.js';

/** Independent malformed/ownership tests derived from our Word-generated
 * fixture. Adding a slot here is test setup, not a production writer or native
 * fidelity claim; the production API must only replace existing operands. */
function setup(variant='normal') {
 const input=new Uint8Array(readFileSync(new URL('./fixtures/doc/rich-runs.doc',import.meta.url)));
 const cfb=unwrapDocBytes(input)!,word=cfb.wordDocBytes.slice(),fib=readDocFib(word),view=new DataView(word.buffer);
 const bte=parseBteTable(cfb.tableBytes,fib.plcfbteChpx),page=bte.pns[0]!*512,count=word[page+511]!,rgb=page+(count+1)*4;
 const target=readDocCharacterRuns(input).find(r=>r.text==='Bold text')!;
 let slot=-1;
 // Independently select the FC interval by its unique bold opcode.
 for(let i=0;i<count;i++){const old=page+word[rgb+i]!*2;for(let at=old+1;at<old+word[old]!;at++)if(word[at]===0x35&&word[at+1]===8)slot=rgb+i;}
 expect(slot).toBeGreaterThanOrEqual(0);
 const old=page+word[slot]!*2,original=word.slice(old+1,old+1+word[old]!),blob=page+200;
 word[slot]=100;word[blob]=original.length+4;word.set(original,blob+1);word.set([0x43,0x4a,36,0],blob+1+original.length);
 const operand=blob+3+original.length;
 if(variant==='shared')word[rgb]=100;
 if(variant==='duplicate'){word[blob]!+=4;word.set([0x43,0x4a,36,0],blob+1+original.length+4);}
 if(variant==='opaque'){word[blob]!+=4;word.set([0x99,0x4a,0,0],blob+1+original.length+4);}
 if(variant==='invalid-existing')view.setUint16(operand,0,true);
 if(variant==='header-alias')word[slot]=2;
 if(variant==='papx-alias'){
  const table=cfb.tableBytes.slice(),papx=parseBteTable(table,fib.plcfbtePapx);
  new DataView(table.buffer).setUint32(fib.plcfbtePapx.fc+(papx.pns.length+1)*4,bte.pns[0]!,true);
  return {input:cfb.rewrap(word,table),target,operand};
 }
 return {input:cfb.rewrap(word,cfb.tableBytes),target,operand};
}
describe('independent DOC font-size operand review',()=>{
 it.each([1,1.5,18,128.5,1638])('changes only the two-byte little-endian operand to %s points',points=>{
  const {input,target,operand}=setup(),before=unwrapDocBytes(input)!,result=writeDocCharacterRunFontSize(input,target.cpStart,target.cpEnd,points),after=unwrapDocBytes(result)!;
  expect(after.wordDocBytes[operand]).toBe(points*2&255);expect(after.wordDocBytes[operand+1]).toBe(points*2>>>8);
  expect(after.tableBytes).toEqual(before.tableBytes);expect(result.length).toBe(input.length);
  const changed=Array.from(after.wordDocBytes.keys()).filter(i=>after.wordDocBytes[i]!==before.wordDocBytes[i]);
  expect(changed.every(i=>i===operand||i===operand+1)).toBe(true);
  expect(readDocCharacterRuns(result).find(r=>r.cpStart===target.cpStart)!.directFontSizePoints).toBe(points);
  expect(readDocCharacterRuns(result).map(r=>r.text)).toEqual(readDocCharacterRuns(input).map(r=>r.text));
 });
 it.each(['shared','duplicate','opaque','invalid-existing','header-alias','papx-alias'])('refuses %s formatting without mutation',variant=>{
  const {input,target}=setup(variant),before=input.slice();expect(()=>writeDocCharacterRunFontSize(input,target.cpStart,target.cpEnd,20)).toThrow();expect(input).toEqual(before);
 });
 it('does not coerce objects, round fractions, overflow or commit a partial paragraph view',()=>{
  const {input,target}=setup();let coercions=0;
  for(const points of [undefined,NaN,Infinity,-1,0.5,1.25,1638.5,65536,{valueOf(){coercions++;return 18;}}])expect(()=>writeDocCharacterRunFontSize(input,target.cpStart,target.cpEnd,points as number)).toThrow('invalid-formatting');
  expect(coercions).toBe(0);expect(()=>writeDocCharacterRunFontSize(input,target.cpStart+1,target.cpEnd,20)).toThrow();
 });
 it('keeps no-op handles current and rejects stale snapshots after any committed edit',()=>{
  const {input}=setup(),doc=new DocDocument(input),run=doc.paragraphs[1]!.runs.find(r=>r.text==='Bold text')!;
  run.directFontSizePoints=18;expect(doc.revision).toBe(0);run.directFontSizePoints=20;
  expect(doc.revision).toBe(1);expect(run.directFontSizePoints).toBe(20);expect(run.sprms.find(p=>p.opcode===0x4a43)!.operand).toEqual([40,0]);
  expect(()=>{run.directFontSizePoints=21;}).toThrow('stale-run');
  const latest=doc.paragraphs[1]!.runs.find(r=>r.text==='Bold text')!;latest.directBold=false;expect(()=>{latest.directFontSizePoints=21;}).toThrow('stale-run');
  expect(doc.paragraphs[1]!.runs.find(r=>r.text==='Bold text')!.directFontSizePoints).toBe(20);
 });
 it('refuses coercive/reentrant numeric objects and preserves the current model',()=>{
  const {input}=setup(),doc=new DocDocument(input),run=doc.paragraphs[1]!.runs.find(r=>r.text==='Bold text')!,before=doc.serialize();let calls=0;
  const value={valueOf(){calls++;run.directBold=false;return 20;}};
  expect(()=>{run.directFontSizePoints=value as unknown as number;}).toThrow(UnsupportedOle2EditError);expect(calls).toBe(0);expect(doc.revision).toBe(0);expect(doc.serialize()).toEqual(before);
 });
});
