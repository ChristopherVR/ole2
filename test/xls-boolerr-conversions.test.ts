import { readFileSync } from 'node:fs';
import { describe, it, expect } from 'vitest';
import { XlsDocument } from '../src/xls-document.js';
import { UnsupportedOle2EditError } from '../src/ole2-document-base.js';
import { readXlsWorkbook } from '../src/legacy-excel-workbook.js';
import { unwrapXlsBytes } from '../src/legacy-excel-cfb.js';
import { readRecords } from '../src/legacy-excel-biff8.js';
import { readPreservedXlsSst } from '../src/legacy-excel-preserved-sst.js';
import { editXlsBoolErrorScalarWorkbookStream } from '../src/legacy-excel-bool-error-conversion.js';
import { editXlsStoredCellWorkbookStream } from '../src/legacy-excel-stored-cell-edit.js';
const fixture = () => new Uint8Array(readFileSync(new URL('./fixtures/xls/workbook-cell-types.xls',import.meta.url)));
const raw = (bytes:Uint8Array) => unwrapXlsBytes(bytes).workbookBytes;

describe('existing BOOLERR scalar value conversions',()=>{
 it.each([1,2,3].flatMap(col=>[37.125,'Converted 漢字 😀'].map(value=>({col,value}))))('converts source column $col to $value preserving every other cell and XF',({col,value})=>{
  const input=fixture(), before=readXlsWorkbook(input), doc=new XlsDocument(input), cell=doc.sheets[0]!.cell(0,col), source=raw(input);
  expect(['boolean','error']).toContain(cell.type);cell.value=value;
  expect(cell.value).toEqual(value);expect(cell.type).toBe(typeof value);
  expect(doc.revision).toBe(1);expect(doc.recalculationRequired).toBe(true);
  const after=readXlsWorkbook(doc.serialize()), selected=after.sheets[0]!.cells.find(c=>c.row===0&&c.col===col)!;
  expect(selected.xf).toBe(before.sheets[0]!.cells.find(c=>c.row===0&&c.col===col)!.xf);
  selected.value=before.sheets[0]!.cells.find(c=>c.row===0&&c.col===col)!.value;expect(after).toEqual(before);
  expect(new XlsDocument(doc.serialize()).sheets[0]!.cell(0,col).value).toEqual(value);
  for(const entry of doc.entries.filter(e=>e.type===2&&e.name!=='Workbook'))expect(doc.getStream(entry.name)).toEqual(new XlsDocument(input).getStream(entry.name));
  const next=raw(doc.serialize()), oldSst=readPreservedXlsSst(source), newSst=readPreservedXlsSst(next);
  expect(newSst.entries.slice(0,oldSst.entries.length).map(e=>({text:e.text,richRuns:e.richRuns,extendedBytes:e.extendedBytes}))).toEqual(oldSst.entries.map(e=>({text:e.text,richRuns:e.richRuns,extendedBytes:e.extendedBytes})));
  expect(newSst.total).toBe(oldSst.total+(typeof value==='string'?1:0));
  const kept=(bytes:Uint8Array)=>{const rs=readRecords(bytes,0,bytes.length),end=rs.find(r=>r.opcode===0xa)!.dataOffset;return rs.filter(r=>r.headerOffset>=end&&![0x205,0x203,0xfd,0x20b,0xd7].includes(r.opcode)).map(r=>bytes.slice(r.headerOffset,r.dataOffset+r.length));};
  expect(kept(next)).toEqual(kept(source));
 });
 it('reuses an existing plain SST entry while leaving its original alias untouched',()=>{
  const input=fixture(), doc=new XlsDocument(input), before=readPreservedXlsSst(raw(input));
  doc.sheets[0]!.cell(0,1).value='Untouched literal';
  expect(doc.sheets[0]!.cell(1,1).value).toBe('Untouched literal');
  const after=readPreservedXlsSst(raw(doc.serialize()));expect(after.entries.length).toBe(before.entries.length);expect(after.total).toBe(before.total+1);
  expect(after.entries.map(e=>e.text)).toEqual(before.entries.map(e=>e.text));
 });
 it('keeps retained handles current through bool/error/number/string edits and no-ops',()=>{
  const doc=new XlsDocument(fixture()), a=doc.sheets[0]!.cell(0,1), b=doc.sheets[0]!.cell(0,2);
  a.value=2.75;b.value='Error resolved';a.value={error:'#REF!'};a.value=8.125;a.value='Ready';
  expect(a.type).toBe('string');expect(b.type).toBe('string');expect(b.value).toBe('Error resolved');expect(doc.revision).toBe(5);
  const bytes=doc.serialize();a.value='Ready';expect(doc.revision).toBe(5);expect(doc.serialize()).toEqual(bytes);
  expect(new XlsDocument(bytes).sheets[0]!.cell(0,1).value).toBe('Ready');
 });
 it('refuses formula caches, clearing, invalid scalar values and missing cells atomically',()=>{
  const input=fixture(),doc=new XlsDocument(input),cell=doc.sheets[0]!.cell(0,1);
  expect(()=>{doc.sheets[0]!.cell(1,3).value=3;}).toThrow(UnsupportedOle2EditError);
  expect(()=>doc.sheets[0]!.cell(50,50)).toThrow(UnsupportedOle2EditError);
  for(const value of [null,NaN,Infinity,-Infinity,'\ud800','x'.repeat(32768)])expect(()=>{cell.value=value as never;}).toThrow(UnsupportedOle2EditError);
  expect(doc.serialize()).toEqual(input);expect(doc.dirty).toBe(false);expect(doc.revision).toBe(0);
 });
 it('strictly validates source Bes discriminators and both defined value sets before changing type',()=>{
  const input=raw(fixture()),records=readRecords(input,0,input.length),target=records.find(r=>r.opcode===0x205)!;
  for(const [flag,value]of [[2,0],[0,2],[1,2],[255,42]]){
   const bytes=input.slice();bytes[target.dataOffset+6]=value!;bytes[target.dataOffset+7]=flag!;
   for(const value of [5,'Invalid source']){const result=editXlsBoolErrorScalarWorkbookStream(bytes,{row:0,col:0,value});expect(result).toMatchObject({status:'unchanged',reason:'malformed-records'});expect(result.bytes).toBe(bytes);}
  }
 });
 it('does not accept wrong source kinds, duplicate coordinates or unsupported future pointers',()=>{
  const input=raw(fixture()),records=readRecords(input,0,input.length),cells=records.filter(r=>r.opcode===0x205);
  const unsupported=editXlsStoredCellWorkbookStream(input,{row:0,col:1,value:2},'invalid' as never);expect(unsupported).toMatchObject({status:'unchanged',reason:'invalid-edit'});expect(unsupported.bytes).toBe(input);
  for(const mutate of [(v:DataView)=>v.setUint16(cells[2]!.dataOffset+2,1,true),(v:DataView)=>v.setUint16(records.find(r=>r.opcode===0x208)!.headerOffset,0x7777,true)]){
   const bytes=input.slice();mutate(new DataView(bytes.buffer));const result=editXlsBoolErrorScalarWorkbookStream(bytes,{row:0,col:1,value:2});expect(result.status).toBe('unchanged');expect(result.bytes).toBe(bytes);
  }
 });
 it('refuses malformed INDEX/DBCELL ownership and shared formula ranges atomically',()=>{
  const input=raw(fixture()),records=readRecords(input,0,input.length),db=records.find(r=>r.opcode===0xd7)!,index=records.find(r=>r.opcode===0x20b)!,row=records.find(r=>r.opcode===0x208)!;
  for(const mutate of [(v:DataView)=>v.setUint32(db.dataOffset,1,true),(v:DataView)=>v.setUint32(index.dataOffset+16,row.headerOffset,true),
   (v:DataView)=>{v.setUint16(row.headerOffset,0x04bc,true);v.setUint16(row.dataOffset,0,true);v.setUint16(row.dataOffset+2,0,true);v.setUint8(row.dataOffset+4,0);v.setUint8(row.dataOffset+5,3);}]){
   const bytes=input.slice();mutate(new DataView(bytes.buffer));const result=editXlsBoolErrorScalarWorkbookStream(bytes,{row:0,col:1,value:'Refused'});expect(result.status).toBe('unchanged');expect(result.bytes).toBe(bytes);
  }
 });
 it('only converts physical BOOLERR records, not ordinary numeric or string sources',()=>{
  const input=raw(fixture());for(const [row,col]of [[1,0],[1,1]]){const result=editXlsBoolErrorScalarWorkbookStream(input,{row:row!,col:col!,value:3});expect(result).toMatchObject({status:'unchanged',reason:'cell-not-supported'});expect(result.bytes).toBe(input);}
 });
});
