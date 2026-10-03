import { readFileSync } from 'node:fs';
import { describe, expect, it } from 'vitest';
import { readPreservedXlsSst, appendXlsSstString } from '../src/legacy-excel-preserved-sst.js';
import { editXlsStringWorkbookStream } from '../src/legacy-excel-preserved-string-edit.js';
import { editXlsPreservedStringCell } from '../src/legacy-excel-preserved-string-cell.js';
import { readXlsWorkbook } from '../src/legacy-excel-workbook.js';
import { unwrapXlsBytes } from '../src/legacy-excel-cfb.js';
import { readRecords } from '../src/legacy-excel-biff8.js';
import { parseOle2 } from '../src/ole2-parser-read.js';
import { readCompoundFileStream } from '../src/ole2-stream-edit.js';

const fixture=(name='workbook-features.xls')=>new Uint8Array(readFileSync(new URL(`./fixtures/xls/${name}`,import.meta.url)));
const raw=(bytes:Uint8Array)=>unwrapXlsBytes(bytes).workbookBytes;
const cell=(bytes:Uint8Array,sheet:number,row:number,col:number)=>readXlsWorkbook(bytes).sheets[sheet]!.cells.find(c=>c.row===row&&c.col===col);

function preservedSst(original:Uint8Array,updated:Uint8Array):void {
 const a=raw(original),b=raw(updated),old=readPreservedXlsSst(a),next=readPreservedXlsSst(b);
 const actual=b.slice(next.start,next.start+(old.end-old.start));
 actual.set(a.subarray(old.start+4,old.start+12),4);
 expect(actual).toEqual(a.subarray(old.start,old.end));
 expect(next.entries.slice(0,old.entries.length).map(e=>[e.text,e.richRuns,e.extendedBytes])).toEqual(old.entries.map(e=>[e.text,e.richRuns,e.extendedBytes]));
}

describe('preserved SST string edits on real Excel fixtures',()=>{
 it.each(['Unicode 日本語 Ω 😀','日本語Ω😀'.repeat(4000)])('appends continued Unicode without rebuilding the cell table',value=>{
  const input=fixture(),before=readXlsWorkbook(input);
  const result=editXlsPreservedStringCell(input,{row:1,col:0,value});
  expect(result.status).toBe('edited');
  expect(cell(result.bytes,0,1,0)?.value).toBe(value);
  preservedSst(input,result.bytes);
  const after=readXlsWorkbook(result.bytes);
  const edited=after.sheets[0]!.cells.find(c=>c.row===1&&c.col===0)!;
  edited.value='Apple';
  expect(after).toEqual(before);
  const originalOle=parseOle2(input.buffer as ArrayBuffer);
  for(const entry of originalOle.entries.filter(e=>e.type===2&&e.name!=='Workbook')) expect(readCompoundFileStream(result.bytes,[entry.name])).toEqual(readCompoundFileStream(input,[entry.name]));
  const old=raw(input),next=raw(result.bytes),oldRecords=readRecords(old,0,old.length),nextRecords=readRecords(next,0,next.length);
  // All worksheet records other than the target/INDEX remain serialized identically.
  const oldEnd=oldRecords.find(r=>r.opcode===0x000a)!.dataOffset;
  const nextEnd=nextRecords.find(r=>r.opcode===0x000a)!.dataOffset;
  const oldSheetRecords=oldRecords.filter(r=>r.headerOffset>=oldEnd&&r.opcode!==0x020b);
  const nextSheetRecords=nextRecords.filter(r=>r.headerOffset>=nextEnd&&r.opcode!==0x020b);
  expect(nextSheetRecords.length).toBe(oldSheetRecords.length);
  oldSheetRecords.forEach((record,i)=>{
   const a=old.subarray(record.headerOffset,record.dataOffset+record.length);
   const nr=nextSheetRecords[i]!;
   const b=next.slice(nr.headerOffset,nr.dataOffset+nr.length);
   if(record.opcode===0xfd&&new DataView(old.buffer).getUint16(record.dataOffset,true)===1&&new DataView(old.buffer).getUint16(record.dataOffset+2,true)===0&&record.headerOffset<oldRecords.filter(r=>r.opcode===0x809)[2]!.headerOffset) b.set(a.subarray(10,14),10);
   expect(b).toEqual(a);
  });
 });
 it('relocates every BOUNDSHEET and INDEX pointer to the original referenced record',()=>{
  const input=raw(fixture()),result=editXlsStringWorkbookStream(input,{row:1,col:0,value:'new text'});
  expect(result.status).toBe('edited');
  const out=result.bytes,iv=new DataView(input.buffer,input.byteOffset,input.byteLength),ov=new DataView(out.buffer);
  const old=readRecords(input,0,input.length),next=readRecords(out,0,out.length),byOffset=new Map(next.map(r=>[r.headerOffset,r]));
  const oldBounds=old.filter(r=>r.opcode===0x85),newBounds=next.filter(r=>r.opcode===0x85);
  oldBounds.forEach((r,i)=>expect(byOffset.get(ov.getUint32(newBounds[i]!.dataOffset,true))?.opcode).toBe(0x809));
  const oldIndexes=old.filter(r=>r.opcode===0x20b),newIndexes=next.filter(r=>r.opcode===0x20b);
  oldIndexes.forEach((r,i)=>{
   for(let p=12;p<r.length;p+=4) {
    const pointer=ov.getUint32(newIndexes[i]!.dataOffset+p,true);
    expect(byOffset.get(pointer)?.opcode).toBe(p===12?0x55:0xd7);
    expect(pointer-iv.getUint32(r.dataOffset+p,true)).toBe(out.length-input.length);
   }
  });
  const table=readPreservedXlsSst(out),ext=next.find(r=>r.opcode===0xff)!,interval=ov.getUint16(ext.dataOffset,true);
  for(let i=0;i<(ext.length-2)/8;i++){
   const entry=table.entries[i*interval]!;
   expect(ov.getUint32(ext.dataOffset+2+i*8,true)).toBe(entry.offset);
   expect(ov.getUint16(ext.dataOffset+6+i*8,true)).toBe(entry.offset-entry.recordOffset);
  }
 });
 it('supports successive edits without removing old SST entries',()=>{
  const input=fixture();
  const first=editXlsPreservedStringCell(input,{row:1,col:0,value:'One Ω'});
  const second=editXlsPreservedStringCell(first.bytes,{row:1,col:0,value:'Two 😀'});
  expect(second.status).toBe('edited');
  expect(cell(second.bytes,0,1,0)?.value).toBe('Two 😀');
  expect(readPreservedXlsSst(raw(second.bytes)).entries.some(e=>e.text==='One Ω')).toBe(true);
  expect(readPreservedXlsSst(raw(second.bytes)).total).toBe(readPreservedXlsSst(raw(input)).total);
 });
 it('refuses unsupported cell types, unknown relocation records and malformed SST counts',()=>{
  const input=raw(fixture());
  expect(editXlsStringWorkbookStream(input,{row:1,col:1,value:'MULRK unsupported'})).toMatchObject({status:'unchanged',reason:'cell-not-supported',bytes:input});
  const unknown=input.slice(),r=readRecords(unknown,0,unknown.length).find(r=>r.opcode===0x0040)!;
  new DataView(unknown.buffer).setUint16(r.headerOffset,0x7777,true);
  expect(editXlsStringWorkbookStream(unknown,{row:1,col:0,value:'new'})).toMatchObject({status:'unchanged',reason:'unsupported-pointer-record',bytes:unknown});
  const malformed=input.slice(),table=readPreservedXlsSst(malformed);
  new DataView(malformed.buffer).setUint32(table.start+8,0x7fffffff,true);
  expect(editXlsStringWorkbookStream(malformed,{row:1,col:0,value:'new'})).toMatchObject({status:'unchanged',reason:'malformed-records',bytes:malformed});
  expect(editXlsStringWorkbookStream(input.subarray(0,input.length-5000),{row:1,col:0,value:'new'}).status).toBe('unchanged');
 });
 it('rejects malformed Unicode and invalid coordinates',()=>{
  const input=fixture();
  for(const edit of [{row:1,col:0,value:'\ud800'},{row:1,col:0,value:'\udc00'},{row:1,col:0,value:'x'.repeat(32768)},{row:-1,col:0,value:'new'}]) expect(editXlsPreservedStringCell(input,edit)).toMatchObject({status:'unchanged',reason:'invalid-edit',bytes:input});
 });
});

describe('strict continuation model',()=>{
 it('keeps each appended record within the BIFF8 maximum and decodes wide/narrow/empty strings',()=>{
  const original=raw(fixture()),table=readPreservedXlsSst(original);
  for(const text of ['', 'a'.repeat(32767),'Ω😀'.repeat(10000)]){
   const appended=appendXlsSstString(text);
   expect(readRecords(appended,0,appended.length).every(r=>r.length<=8224)).toBe(true);
   const out=new Uint8Array([...original.subarray(0,table.end),...appended,...original.subarray(table.end)]);
   new DataView(out.buffer).setUint32(table.start+8,table.entries.length+1,true);
   expect(readPreservedXlsSst(out).entries.at(-1)?.text).toBe(text);
  }
 });
});

describe('rich strings and aliases from native Excel',()=>{
 const name='workbook-rich-strings.xls';
 it('preserves a shared alias and unrelated rich entry including raw formatting runs',()=>{
  const input=fixture(name),source=raw(input),sst=readPreservedXlsSst(source);
  expect(sst.entries.some(e=>e.richRuns===3)).toBe(true);
  const labels=readRecords(source,0,source.length).filter(r=>r.opcode===0xfd);
  const v=new DataView(source.buffer,source.byteOffset,source.byteLength);
  const indices=labels.filter(r=>v.getUint16(r.dataOffset,true)===0&&v.getUint16(r.dataOffset+2,true)<=1).slice(0,2).map(r=>v.getUint32(r.dataOffset+6,true));
  expect(indices[0]).toBe(indices[1]);
  const result=editXlsPreservedStringCell(input,{row:0,col:0,value:'New shared target 日本語 😀'});
  expect(result.status).toBe('edited');
  expect(cell(result.bytes,0,0,0)?.value).toBe('New shared target 日本語 😀');
  expect(cell(result.bytes,0,0,1)?.value).toBe('Shared plain text');
  expect(cell(result.bytes,0,1,0)?.value).toBe('Rich preserved Ω');
  preservedSst(input,result.bytes);
 });
 it('replaces a rich target with plain text without changing its shared rich entry',()=>{
  const input=fixture(name),before=readPreservedXlsSst(raw(input));
  const result=editXlsPreservedStringCell(input,{row:1,col:0,value:'Rich preserved Ω'});
  expect(result.status).toBe('edited');
  expect(result.bytes.length).toBe(input.length);
  const bytes=raw(result.bytes),labels=readRecords(bytes,0,bytes.length),view=new DataView(bytes.buffer,bytes.byteOffset,bytes.byteLength);
  const target=labels.find(r=>r.opcode===0xfd&&view.getUint16(r.dataOffset,true)===1&&view.getUint16(r.dataOffset+2,true)===0)!;
  const after=readPreservedXlsSst(bytes),index=view.getUint32(target.dataOffset+6,true);
  expect(after.entries[index]?.richRuns).toBe(0);
  expect(after.entries.map(e=>[e.text,e.richRuns,e.extendedBytes])).toEqual(before.entries.map(e=>[e.text,e.richRuns,e.extendedBytes]));
 });
 it('edits a later tab and converts an existing RK without rebuilding rows',()=>{
  const input=fixture(name);
  const result=editXlsPreservedStringCell(input,{worksheetIndex:1,row:1,col:0,value:'A number becomes text Ω'});
  expect(result.status).toBe('edited');
  expect(cell(result.bytes,1,1,0)?.value).toBe('A number becomes text Ω');
  expect(readXlsWorkbook(result.bytes).sheets[0]).toEqual(readXlsWorkbook(input).sheets[0]);
  expect(readPreservedXlsSst(raw(result.bytes)).total).toBe(readPreservedXlsSst(raw(input)).total+1);
 });
 it('honors reordered worksheet tab directory and refuses duplicate target records',()=>{
  const source=raw(fixture(name)),view=new DataView(source.buffer,source.byteOffset,source.byteLength);
  const bounds=readRecords(source,0,source.length).filter(r=>r.opcode===0x85);
  // Reorder tab associations by swapping both complete BOUNDSHEET records.
  const originalRecords=readRecords(source,0,source.length);
  const first=source.subarray(bounds[0]!.headerOffset,bounds[0]!.dataOffset+bounds[0]!.length);
  const second=source.subarray(bounds[1]!.headerOffset,bounds[1]!.dataOffset+bounds[1]!.length);
  const reordered=new Uint8Array([...source.subarray(0,bounds[0]!.headerOffset),...second,...source.subarray(bounds[0]!.dataOffset+bounds[0]!.length,bounds[1]!.headerOffset),...first,...source.subarray(bounds[1]!.dataOffset+bounds[1]!.length)]);
  const result=editXlsStringWorkbookStream(reordered,{worksheetIndex:0,row:0,col:0,value:'Reordered later Ω'});
  expect(result.status).toBe('edited');
  expect(readXlsWorkbook(result.bytes).sheets[0]?.name).toBe('Later');
  expect(cell(result.bytes,0,0,0)?.value).toBe('Reordered later Ω');
  expect(cell(result.bytes,1,0,0)?.value).toBe('Shared plain text');
  const duplicate=source.slice();
  const targets=originalRecords.filter(r=>r.opcode===0xfd&&view.getUint16(r.dataOffset,true)===0&&view.getUint16(r.dataOffset+2,true)<=1).slice(0,2);
  new DataView(duplicate.buffer).setUint16(targets[1]!.dataOffset+2,0,true);
  expect(editXlsStringWorkbookStream(duplicate,{row:0,col:0,value:'new'})).toMatchObject({status:'unchanged',reason:'ambiguous-cell',bytes:duplicate});
 });
});

function physical(opcode:number,data:number[]):number[]{return[opcode&255,opcode>>>8,data.length&255,data.length>>>8,...data];}
function sstChain(parts:number[][]):Uint8Array {return new Uint8Array([...physical(0xfc,[1,0,0,0,1,0,0,0,...parts[0]!]),...parts.slice(1).flatMap(p=>physical(0x3c,p)),...physical(0xa,[])]);}
describe('malformed continuation bounds and rich/ext boundaries',()=>{
 it('checks character flags and never reads a split UTF16 code unit into the next record',()=>{
  expect(readPreservedXlsSst(sstChain([[2,0,1,65,0],[1,66,0]])).entries[0]?.text).toBe('AB');
  expect(()=>readPreservedXlsSst(sstChain([[2,0,1,65,0],[2,66,0]]))).toThrow();
  expect(()=>readPreservedXlsSst(sstChain([[2,0,1,65],[1,0,66,0]]))).toThrow();
  expect(()=>readPreservedXlsSst(sstChain([[2,0,1,65,0],[1,66]]))).toThrow();
 });
 it('does not interpret rich/extended metadata continuation bytes as character flags',()=>{
  const rich=sstChain([[1,0,8,1,0,65,0,0],[2,0]]);
  expect(readPreservedXlsSst(rich).entries[0]).toMatchObject({text:'A',richRuns:1});
  const ext=sstChain([[1,0,4,4,0,0,0,65,99],[88,77,66]]);
  expect(readPreservedXlsSst(ext).entries[0]).toMatchObject({text:'A',extendedBytes:4});
  expect(()=>readPreservedXlsSst(sstChain([[1,0,8,1,0,65,0,0],[2]]))).toThrow();
 });
});
