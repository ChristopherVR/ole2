import{readFileSync}from'node:fs';
import{describe,it,expect}from'vitest';
import{editXlsPreservedStringCell}from'../src/legacy-excel-preserved-string-cell.js';
import{editXlsStringWorkbookStream}from'../src/legacy-excel-preserved-string-edit.js';
import{readXlsWorkbook}from'../src/legacy-excel-workbook.js';
import{readPreservedXlsSst}from'../src/legacy-excel-preserved-sst.js';
import{unwrapXlsBytes}from'../src/legacy-excel-cfb.js';
import{readRecords}from'../src/legacy-excel-biff8.js';
import{replaceXlsNumericRecord}from'../src/legacy-excel-cell-record-splices.js';
const fixture=()=>new Uint8Array(readFileSync(new URL('./fixtures/xls/workbook-features.xls',import.meta.url)));
const raw=(b:Uint8Array)=>unwrapXlsBytes(b).workbookBytes;
const cell=(b:Uint8Array,row:number,col:number)=>readXlsWorkbook(b).sheets[0]!.cells.find(c=>c.row===row&&c.col===col);
const values=(b:Uint8Array)=>readXlsWorkbook(b);

describe('preserved numeric-record to string conversions',()=>{
 it.each([{row:1,col:1,opcode:0x00bd},{row:1,col:2,opcode:0x00bd},{row:4,col:1,opcode:0x0203}])('converts $opcode at [$row,$col] while preserving every other modeled cell',({row,col,opcode})=>{
  const input=fixture(),source=raw(input),view=new DataView(source.buffer,source.byteOffset,source.byteLength);
  expect(readRecords(source,0,source.length).some(r=>r.opcode===opcode&&view.getUint16(r.dataOffset,true)===row)).toBe(true);
  const before=values(input),result=editXlsPreservedStringCell(input,{row,col,value:'Converted 日本語 😀'});
  expect(result.status).toBe('edited');
  expect(cell(result.bytes,row,col)?.value).toBe('Converted 日本語 😀');
  expect(cell(result.bytes,row,col)?.xf).toBe(cell(input,row,col)?.xf);
  const after=values(result.bytes),target=after.sheets[0]!.cells.find(c=>c.row===row&&c.col===col)!;
  target.value=before.sheets[0]!.cells.find(c=>c.row===row&&c.col===col)!.value;
  expect(after).toEqual(before);
  const oldSst=readPreservedXlsSst(source),next=raw(result.bytes),nextSst=readPreservedXlsSst(next);
  const preserved=next.slice(nextSst.start,nextSst.start+oldSst.end-oldSst.start);
  preserved.set(source.subarray(oldSst.start+4,oldSst.start+12),4);
  expect(preserved).toEqual(source.subarray(oldSst.start,oldSst.end));
  expect(nextSst.total).toBe(oldSst.total+1);
  // Cell-table bytes remain identical after excluding only replaced numeric
  // records and INDEX/DBCELL address fields; formula/ROW/drawing records survive.
  const oldRecords=readRecords(source,0,source.length),newRecords=readRecords(next,0,next.length);
  const globalsEnd=oldRecords.find(r=>r.opcode===0xa)!.dataOffset;
  const nextGlobalsEnd=newRecords.find(r=>r.opcode===0xa)!.dataOffset;
  const preservedOps=(rs:typeof oldRecords,end:number)=>rs.filter(r=>r.headerOffset>=end&&!new Set([0x203,0x27e,0xbd,0xfd,0x20b,0xd7]).has(r.opcode));
  const a=preservedOps(oldRecords,globalsEnd),b=preservedOps(newRecords,nextGlobalsEnd);
  expect(a.length).toBe(b.length);
  a.forEach((r,i)=>expect(next.subarray(b[i]!.headerOffset,b[i]!.dataOffset+b[i]!.length)).toEqual(source.subarray(r.headerOffset,r.dataOffset+r.length)));
 });
 it.each([{row:1,col:1},{row:1,col:2},{row:4,col:1}])('correctly relocates row-block and sheet pointers for [$row,$col]',({row,col})=>{
  const input=raw(fixture()),result=editXlsStringWorkbookStream(input,{row,col,value:'Name'});
  expect(result.status).toBe('edited');
  const next=result.bytes,view=new DataView(next.buffer),records=readRecords(next,0,next.length),byOffset=new Map(records.map(r=>[r.headerOffset,r]));
  for(const db of records.filter(r=>r.opcode===0xd7)){
   const back=view.getUint32(db.dataOffset,true);
   if(back===0)continue;
   const rowRecord=byOffset.get(db.headerOffset-back)!;
   expect(rowRecord?.opcode).toBe(0x208);
   const rows=records.filter(r=>r.opcode===0x208&&r.headerOffset>=rowRecord.headerOffset&&r.headerOffset<db.headerOffset);
   expect(rows.length).toBe((db.length-4)/2);
   let base=rowRecord.dataOffset+rowRecord.length;
   for(let i=0;i<rows.length;i++){
    const address=base+view.getUint16(db.dataOffset+4+i*2,true),first=byOffset.get(address)!;
    expect(first).toBeDefined();
    expect(view.getUint16(first.dataOffset,true)).toBe(view.getUint16(rows[i]!.dataOffset,true));
    base=address;
   }
  }
  for(const index of records.filter(r=>r.opcode===0x20b)) for(let p=12;p<index.length;p+=4) expect(byOffset.get(view.getUint32(index.dataOffset+p,true))?.opcode).toBe(p===12?0x55:0xd7);
  for(const bound of records.filter(r=>r.opcode===0x85)) expect(byOffset.get(view.getUint32(bound.dataOffset,true))?.opcode).toBe(0x809);
 });
 it('fails atomically for invalid DBCELL pointers and shared formula ranges',()=>{
  const input=raw(fixture()),records=readRecords(input,0,input.length);
  const db=records.find(r=>r.opcode===0xd7)!;
  for(const mutate of [(v:DataView)=>v.setUint32(db.dataOffset,1,true),(v:DataView)=>v.setUint16(db.dataOffset+4,65535,true)]){
   const malformed=input.slice();mutate(new DataView(malformed.buffer));
   expect(editXlsStringWorkbookStream(malformed,{row:1,col:1,value:'Name'})).toMatchObject({status:'unchanged',reason:'malformed-records',bytes:malformed});
  }
  // Reinterpret an existing ROW-sized record as a shared-formula definition;
  // its first6bytes cover the target range, without resizing the fixture.
  const definition=records.find(r=>r.opcode===0x208)!;
  const shared=input.slice(),sv=new DataView(shared.buffer);
  sv.setUint16(definition.headerOffset,0x04bc,true);
  sv.setUint16(definition.dataOffset,1,true);sv.setUint16(definition.dataOffset+2,1,true);
  sv.setUint8(definition.dataOffset+4,1);sv.setUint8(definition.dataOffset+5,2);
  expect(editXlsStringWorkbookStream(shared,{row:1,col:1,value:'Name'})).toMatchObject({status:'unchanged',reason:'cell-not-supported',bytes:shared});
 });
});

function record(opcode:number,data:number[]):Uint8Array{return new Uint8Array([opcode&255,opcode>>>8,data.length&255,data.length>>>8,...data]);}
describe('MULRK raw sibling preservation',()=>{
 it.each([0,1,2,3,4])('splits selected cell %i preserving exact sibling ixfe/RK bytes',selected=>{
  const parts=Array.from({length:5},(_,i)=>[15+i,0,((i+1)<<2)|2,0,0,0]);
  const input=record(0xbd,[7,0,10,0,...parts.flat(),14,0]);
  const target=readRecords(input,0,input.length)[0]!,result=replaceXlsNumericRecord(input,target,10+selected,3);
  const out=result.bytes,view=new DataView(out.buffer);
  const seen=new Map<number,number[]>();
  for(const r of readRecords(out,0,out.length)){
   if(r.opcode===0xfd){expect(view.getUint16(r.dataOffset+2,true)).toBe(10+selected);expect(view.getUint16(r.dataOffset+4,true)).toBe(15+selected);expect(view.getUint32(r.dataOffset+6,true)).toBe(3);continue;}
   const first=view.getUint16(r.dataOffset+2,true),count=r.opcode===0xbd?(r.length-6)/6:1;
   for(let i=0;i<count;i++) seen.set(first+i,Array.from(out.subarray(r.dataOffset+4+i*6,r.dataOffset+10+i*6)));
  }
  expect(seen.size).toBe(4);
  parts.forEach((p,i)=>{if(i!==selected)expect(seen.get(10+i)).toEqual(p);});
 });
});

describe('native packed numeric group with distinct formats',()=>{
 const input=new Uint8Array(readFileSync(new URL('./fixtures/xls/workbook-mulrk.xls',import.meta.url)));
 it.each([0,1,2,3])('converts packed cell %i while preserving sibling values/formats and aliases',col=>{
  const before=readXlsWorkbook(input),result=editXlsPreservedStringCell(input,{row:0,col,value:'Numeric conversion 日本語 😀'});
  expect(result.status).toBe('edited');
  const after=readXlsWorkbook(result.bytes),target=after.sheets[0]!.cells.find(c=>c.row===0&&c.col===col)!;
  expect(target.value).toBe('Numeric conversion 日本語 😀');
  const original=before.sheets[0]!.cells.find(c=>c.row===0&&c.col===col)!;
  expect(target.xf).toBe(original.xf);
  target.value=original.value;
  expect(after).toEqual(before);
  expect(after.sheets[0]!.cells.filter(c=>c.row===0).map(c=>c.xf)).toEqual([62,63,64,65]);
  expect(after.sheets[0]!.cells.filter(c=>c.row===1&&c.col<2).map(c=>c.value)).toEqual(['Alias anchor','Alias anchor']);
 });
 it('does not modify any unselected numeric subrecord bytes',()=>{
  const source=raw(input),sv=new DataView(source.buffer,source.byteOffset,source.byteLength);
  const packed=readRecords(source,0,source.length).find(r=>r.opcode===0xbd&&sv.getUint16(r.dataOffset,true)===0)!;
  const result=editXlsPreservedStringCell(input,{row:0,col:1,value:'Alias anchor'});
  expect(result.status).toBe('edited');
  const next=raw(result.bytes),view=new DataView(next.buffer,next.byteOffset,next.byteLength);
  const numerics=readRecords(next,0,next.length).filter(r=>[0x27e,0xbd].includes(r.opcode)&&view.getUint16(r.dataOffset,true)===0);
  for(const r of numerics){
   const first=view.getUint16(r.dataOffset+2,true),count=r.opcode===0xbd?(r.length-6)/6:1;
   for(let i=0;i<count;i++)expect(next.subarray(r.dataOffset+4+i*6,r.dataOffset+10+i*6)).toEqual(source.subarray(packed.dataOffset+4+(first+i)*6,packed.dataOffset+10+(first+i)*6));
  }
  expect(readPreservedXlsSst(next).entries.length).toBe(readPreservedXlsSst(source).entries.length);
 });
});

it('refuses invalid direct numeric-record replacement arguments instead of wrapping bytes',()=>{
 const input=record(0xbd,[7,0,10,0,15,0,42,0,0,0,16,0,82,0,0,0,11,0]);
 const target=readRecords(input,0,input.length)[0]!;
 for(const [col,index]of[[9,0],[12,0],[10,-1],[10,0x80000000],[10,0.5],[10,NaN]]) expect(()=>replaceXlsNumericRecord(input,target,col,index)).toThrow();
 for(const malformed of [{...target,opcode:0x7777},{...target,headerOffset:-1},{...target,dataOffset:0},{...target,length:target.length+2}]) expect(()=>replaceXlsNumericRecord(input,malformed,10,0)).toThrow();
 const broken=input.slice();new DataView(broken.buffer).setUint16(broken.length-2,255,true);
 expect(()=>replaceXlsNumericRecord(broken,target,10,0)).toThrow();
});
