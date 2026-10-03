import { readFileSync } from 'node:fs';
import { describe, it, expect } from 'vitest';
import { XlsDocument } from '../src/xls-document.js';
import { UnsupportedOle2EditError } from '../src/ole2-document-base.js';
import { readXlsWorkbook, type XlsErrorCode } from '../src/legacy-excel-workbook.js';
import { unwrapXlsBytes } from '../src/legacy-excel-cfb.js';
import { readRecords } from '../src/legacy-excel-biff8.js';
import { editXlsBoolErrorCell, editXlsBoolErrorWorkbookStream } from '../src/legacy-excel-bool-error-edit.js';
const fixture = (name='workbook-mulrk.xls') => new Uint8Array(readFileSync(new URL(`./fixtures/xls/${name}`,import.meta.url)));
const raw = (b:Uint8Array) => unwrapXlsBytes(b).workbookBytes;

describe('typed boolean and error cell writes', () => {
 it.each([0,1,2,3])('splits packed column %i preserving every other cell and formatting', col => {
  const input=fixture(),before=readXlsWorkbook(input),document=new XlsDocument(input),cell=document.sheets[0]!.cell(0,col);
  expect(cell.type).toBe('number'); cell.value=true;
  expect(cell.type).toBe('boolean');expect(cell.value).toBe(true);
  const after=readXlsWorkbook(document.serialize()),selected=after.sheets[0]!.cells.find(c=>c.row===0&&c.col===col)!;
  selected.value=before.sheets[0]!.cells.find(c=>c.row===0&&c.col===col)!.value;
  expect(after).toEqual(before);
  const source=raw(input),out=raw(document.serialize());
  const oldGlobals=readRecords(source,0,source.length).find(r=>r.opcode===0xa)!.dataOffset;
  expect(out.subarray(0,oldGlobals)).toEqual(source.subarray(0,oldGlobals));
  // Globals/SST precede the selected row and retain exact bytes. Later tab
  // offsets may move in workbooks with multiple sheets, covered separately.
  expect(document.recalculationRequired).toBe(true);
  const untouched=(bytes:Uint8Array)=>readRecords(bytes,0,bytes.length).filter(r=>![0x203,0x27e,0xbd,0x205,0x20b,0xd7,0x85].includes(r.opcode)).map(r=>bytes.slice(r.headerOffset,r.dataOffset+r.length));
  expect(untouched(out)).toEqual(untouched(source));
 });
 it.each(['#NULL!','#DIV/0!','#VALUE!','#REF!','#NAME?','#NUM!','#N/A','#GETTING_DATA'] as XlsErrorCode[])('stores exact native Bes error %s',error=>{
  const document=new XlsDocument(fixture()),cell=document.sheets[0]!.cell(0,1);
  cell.value={error};expect(cell.type).toBe('error');expect(cell.value).toEqual({error});
  expect(new XlsDocument(document.serialize()).sheets[0]!.cell(0,1).value).toEqual({error});
 });
 it.each([{row:1,col:1},{row:1,col:2},{row:4,col:1}])('converts real NUMBER/RK/MULRK at [$row,$col] with all pointers readable',({row,col})=>{
  const input=fixture('workbook-features.xls'),before=readXlsWorkbook(input),document=new XlsDocument(input);
  document.sheets[0]!.cell(row,col).value={error:'#DIV/0!'};
  const after=readXlsWorkbook(document.serialize());
  after.sheets[0]!.cells.find(c=>c.row===row&&c.col===col)!.value=before.sheets[0]!.cells.find(c=>c.row===row&&c.col===col)!.value;
  expect(after).toEqual(before);
  const out=raw(document.serialize()),view=new DataView(out.buffer,out.byteOffset,out.byteLength),records=readRecords(out,0,out.length),byOffset=new Map(records.map(r=>[r.headerOffset,r]));
  for(const r of records.filter(r=>r.opcode===0x85))expect(byOffset.get(view.getUint32(r.dataOffset,true))?.opcode).toBe(0x809);
  for(const r of records.filter(r=>r.opcode===0x20b))for(let p=r.dataOffset+12;p<r.dataOffset+r.length;p+=4)expect(byOffset.get(view.getUint32(p,true))?.opcode).toBe(p===r.dataOffset+12?0x55:0xd7);
  for(const db of records.filter(r=>r.opcode===0xd7)){
   const back=view.getUint32(db.dataOffset,true);if(!back)continue;
   const firstRow=byOffset.get(db.headerOffset-back)!;expect(firstRow.opcode).toBe(0x208);
   const rows=records.filter(r=>r.opcode===0x208&&r.headerOffset>=firstRow.headerOffset&&r.headerOffset<db.headerOffset);
   expect(rows.length).toBe((db.length-4)/2);let base=firstRow.dataOffset+firstRow.length;
   for(let i=0;i<rows.length;i++){
    const address=base+view.getUint16(db.dataOffset+4+i*2,true),first=byOffset.get(address)!;
    expect(first).toBeDefined();expect(view.getUint16(first.dataOffset,true)).toBe(view.getUint16(rows[i]!.dataOffset,true));base=address;
   }
  }
 });
 it('keeps handles current through boolean/error toggles and later edits, with detached error objects',()=>{
  const document=new XlsDocument(fixture()),a=document.sheets[0]!.cell(0,1),b=document.sheets[0]!.cell(0,3);
  a.value=true;b.value=false;a.value={error:'#N/A'};
  expect(a.value).toEqual({error:'#N/A'});expect(a.type).toBe('error');expect(b.value).toBe(false);
  const error=a.value as {error:string};error.error='tamper';expect(a.value).toEqual({error:'#N/A'});
  const revision=document.revision;a.value={error:'#N/A'};expect(document.revision).toBe(revision);
  a.value=false;expect(a.value).toBe(false);expect(a.type).toBe('boolean');
  const unchangedRevision=document.revision;
  a.value=2;expect(a.type).toBe('number');expect(a.value).toBe(2);
  a.value='text';expect(a.type).toBe('string');expect(a.value).toBe('text');
  expect(document.revision).toBe(unchangedRevision+2);
 });
 it('rejects string/formula targets and invalid errors without dirtying the document',()=>{
  const bytes=fixture(),document=new XlsDocument(bytes),sheet=document.sheets[0]!;
  expect(()=>{sheet.cell(1,0).value=true;}).toThrow(UnsupportedOle2EditError);
  expect(sheet.cell(3,0).type).toBe('formula');
  expect(()=>{sheet.cell(3,0).value={error:'#N/A'};}).toThrow(UnsupportedOle2EditError);
  expect(()=>{sheet.cell(0,1).value={error:'invalid'} as never;}).toThrow(UnsupportedOle2EditError);
  expect(()=>{sheet.cell(0,1).value={error:'#N/A',extra:1} as never;}).toThrow(UnsupportedOle2EditError);
  expect(document.serialize()).toEqual(bytes);expect(document.dirty).toBe(false);
 });
 it('rejects malformed original BOOLERR flags and codes and truncated framing atomically',()=>{
  const initial=editXlsBoolErrorWorkbookStream(raw(fixture()),{row:0,col:1,value:true});expect(initial.status).toBe('edited');
  const records=readRecords(initial.bytes,0,initial.bytes.length),r=records.find(r=>r.opcode===0x205)!;
  for(const [offset,value]of [[7,2],[6,2]] as const){
   const bytes=initial.bytes.slice();bytes[r.dataOffset+offset]=value;
   const result=editXlsBoolErrorWorkbookStream(bytes,{row:0,col:1,value:false});
   expect(result).toMatchObject({status:'unchanged',reason:'malformed-records'});expect(result.bytes).toBe(bytes);
  }
  const invalid=editXlsBoolErrorCell(fixture(),{row:0,col:1,value:{error:'bad'} as never});expect(invalid.status).toBe('unchanged');
  const truncated=raw(fixture()).subarray(0,60);expect(editXlsBoolErrorWorkbookStream(truncated,{row:0,col:1,value:true})).toMatchObject({status:'unchanged',bytes:truncated});
 });
 it.each([{col:1,value:{error:'#VALUE!'} as const},{col:2,value:false},{col:3,value:{error:'#REF!'} as const}])('edits native BOOLERR column $col without resizing or changing another byte',({col,value})=>{
  const input=fixture('workbook-cell-types.xls'),document=new XlsDocument(input),source=raw(input);
  const records=readRecords(source,0,source.length),view=new DataView(source.buffer,source.byteOffset,source.byteLength);
  const selected=records.find(r=>r.opcode===0x205&&view.getUint16(r.dataOffset,true)===0&&view.getUint16(r.dataOffset+2,true)===col)!;
  expect(selected).toBeDefined();document.sheets[0]!.cell(0,col).value=value;
  const out=raw(document.serialize());expect(out.length).toBe(source.length);
  const restored=out.slice();restored.set(source.subarray(selected.dataOffset+6,selected.dataOffset+8),selected.dataOffset+6);expect(restored).toEqual(source);
  const a=readXlsWorkbook(input),b=readXlsWorkbook(document.serialize());
  b.sheets[0]!.cells.find(c=>c.row===0&&c.col===col)!.value=a.sheets[0]!.cells.find(c=>c.row===0&&c.col===col)!.value;expect(b).toEqual(a);
  expect(document.sheets[0]!.cell(0,col).snapshot.type).toBe(typeof value==='boolean'?'boolean':'error');
 });
 it('refuses unknown resize records and invalid DBCELL addresses atomically',()=>{
  const source=raw(fixture()),records=readRecords(source,0,source.length);
  for(const mode of ['unknown','dbcell']){
   const bytes=source.slice(),view=new DataView(bytes.buffer);
   if(mode==='unknown')view.setUint16(records.find(r=>r.opcode===0x55)!.headerOffset,0xffff,true);
   else view.setUint32(records.find(r=>r.opcode===0xd7)!.dataOffset,1,true);
   const result=editXlsBoolErrorWorkbookStream(bytes,{row:0,col:1,value:true});expect(result.status).toBe('unchanged');expect(result.bytes).toBe(bytes);
  }
 });
 it('captures an accessor-backed error once before checking or serializing it',()=>{
  let reads=0;const value={get error(){reads++;return reads===1?'#REF!':'invalid';}};
  const result=editXlsBoolErrorWorkbookStream(raw(fixture()),{row:0,col:1,value:value as never});
  expect(result.status).toBe('edited');expect(reads).toBe(1);
  expect(readXlsWorkbook(result.bytes).sheets[0]!.cells.find(c=>c.row===0&&c.col===1)!.value).toEqual({error:'#REF!'});
  reads=0;const document=new XlsDocument(fixture());document.sheets[0]!.cell(0,1).value=value as never;
  expect(reads).toBe(1);expect(document.sheets[0]!.cell(0,1).value).toEqual({error:'#REF!'});
 });
});
