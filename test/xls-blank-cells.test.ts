import { readFileSync } from 'node:fs';
import { describe, it, expect } from 'vitest';
import { XlsDocument } from '../src/xls-document.js';
import { UnsupportedOle2EditError } from '../src/ole2-document-base.js';
import { readXlsWorkbook } from '../src/legacy-excel-workbook.js';
import { unwrapXlsBytes } from '../src/legacy-excel-cfb.js';
import { readRecords } from '../src/legacy-excel-biff8.js';
import { editXlsBlankWorkbookStream } from '../src/legacy-excel-blank-edit.js';
import { replaceXlsBlankRecord } from '../src/legacy-excel-blank-record.js';
const fixture = () => new Uint8Array(readFileSync(new URL('./fixtures/xls/workbook-blanks.xls', import.meta.url)));
const raw = (bytes: Uint8Array) => unwrapXlsBytes(bytes).workbookBytes;

describe('stored blank cell model values', () => {
 it.each([0,1,2,3])('fills packed column %i with NUMBER while preserving every sibling and XF', col => {
  const input = fixture(), before = readXlsWorkbook(input), doc = new XlsDocument(input), cell = doc.sheets[0]!.cell(0,col);
  expect(cell.type).toBe('blank'); cell.value = 12.75;
  expect(cell.type).toBe('number'); expect(cell.value).toBe(12.75);
  const after = readXlsWorkbook(doc.serialize()), selected = after.sheets[0]!.cells.find(c=>c.row===0&&c.col===col)!;
  expect(selected.xf).toBe(before.sheets[0]!.cells.find(c=>c.row===0&&c.col===col)!.xf);
  selected.value = null; expect(after).toEqual(before);
  expect(doc.recalculationRequired).toBe(true); expect(doc.revision).toBe(1);
 });
 it.each([37.125,'Blank Unicode 漢字 😀',true,{error:'#N/A'} as const])('fills single BLANK with %j transactionally', value => {
  const input=fixture(), before=readXlsWorkbook(input), doc=new XlsDocument(input), cell=doc.sheets[0]!.cell(1,1);
  cell.value=value; expect(cell.value).toEqual(value);
  expect(new XlsDocument(doc.serialize()).sheets[0]!.cell(1,1).value).toEqual(value);
  const after=readXlsWorkbook(doc.serialize()); after.sheets[0]!.cells.find(c=>c.row===1&&c.col===1)!.value=null;
  expect(after).toEqual(before);
  for(const entry of doc.entries.filter(e=>e.type===2&&e.name!=='Workbook')) expect(doc.getStream(entry.name)).toEqual(new XlsDocument(input).getStream(entry.name));
 });
 it('keeps retained handles current across multiple type-changing edits and detached snapshots',()=>{
  const doc=new XlsDocument(fixture()), a=doc.sheets[0]!.cell(0,0), b=doc.sheets[0]!.cell(0,1), later=doc.sheets[1]!.cell(1,1);
  a.value='First'; b.value=true; later.value={error:'#DIV/0!'};
  expect(a.value).toBe('First'); expect(b.value).toBe(true); expect(later.type).toBe('error');
  expect(a.xf).toBe(62); expect(b.xf).toBe(63); expect(doc.revision).toBe(3);
  const value=later.value as {error:string}; value.error='tamper'; expect(later.value).toEqual({error:'#DIV/0!'});
  expect(new XlsDocument(doc.serialize()).sheets[0]!.cell(0,0).value).toBe('First');
 });
 it('preserves pristine state for blank no-op, missing, merged interior, invalid and reentrant writes',()=>{
  const input=fixture(), doc=new XlsDocument(input), cell=doc.sheets[0]!.cell(1,1);
  cell.value=null; expect(doc.dirty).toBe(false);
  expect(()=>doc.sheets[0]!.cell(40,40)).toThrow(UnsupportedOle2EditError);
  expect(()=>{doc.sheets[0]!.cell(4,3).value=1;}).toThrow(UnsupportedOle2EditError);
  for(const value of [NaN,Infinity,{},'\ud800']) expect(()=>{cell.value=value as never;}).toThrow(UnsupportedOle2EditError);
  const hostile={get error(){doc.sheets[0]!.cell(0,0).value=3; return '#N/A';}};
  expect(()=>{cell.value=hostile as never;}).toThrow(UnsupportedOle2EditError);
  expect(doc.serialize()).toEqual(input); expect(doc.revision).toBe(0); expect(doc.recalculationRequired).toBe(false);
  cell.value=1; expect(cell.value).toBe(1); // Guard releases after failures.
 });
 it('captures error accessors once, rejecting an invalid captured value without mutation',()=>{
  const doc=new XlsDocument(fixture()), cell=doc.sheets[0]!.cell(1,1);let reads=0;
  cell.value={get error(){reads++;return reads===1?'#N/A':'invalid';}} as never;
  expect(reads).toBe(1);expect(cell.value).toEqual({error:'#N/A'});
 });
 it('refuses ambiguous and malformed packed physical cells without discarding original bytes',()=>{
  const input=raw(fixture()), records=readRecords(input,0,input.length), packed=records.find(r=>r.opcode===0xbe)!, blank=records.find(r=>r.opcode===0x201)!;
  const mutations=[(v:DataView)=>v.setUint16(packed.dataOffset+packed.length-2,255,true),
   (v:DataView)=>v.setUint16(packed.dataOffset+2,256,true),
   (v:DataView)=>{v.setUint16(blank.dataOffset,0,true);v.setUint16(blank.dataOffset+2,1,true);},
   (v:DataView)=>v.setUint16(packed.headerOffset,0x7777,true)];
  for(const mutate of mutations){const bytes=input.slice();mutate(new DataView(bytes.buffer));const r=editXlsBlankWorkbookStream(bytes,{row:0,col:1,value:2});expect(r.status).toBe('unchanged');expect(r.bytes).toBe(bytes);}
 });
 it('refuses malformed DBCELL and structural formula overlap atomically',()=>{
  const input=raw(fixture()), records=readRecords(input,0,input.length), db=records.find(r=>r.opcode===0xd7)!, row=records.find(r=>r.opcode===0x208)!;
  for(const change of [(v:DataView)=>v.setUint32(db.dataOffset,1,true),
   (v:DataView)=>{v.setUint16(row.headerOffset,0x04bc,true);v.setUint16(row.dataOffset,0,true);v.setUint16(row.dataOffset+2,0,true);v.setUint8(row.dataOffset+4,0);v.setUint8(row.dataOffset+5,3);}]){
   const bytes=input.slice();change(new DataView(bytes.buffer));const result=editXlsBlankWorkbookStream(bytes,{row:0,col:1,value:3});expect(result.status).toBe('unchanged');expect(result.bytes).toBe(bytes);
  }
 });
 it('refuses cross-sheet INDEX pointer aliases rather than relocating foreign row blocks',()=>{
  const source=raw(fixture()), records=readRecords(source,0,source.length), indices=records.filter(r=>r.opcode===0x20b), dbs=records.filter(r=>r.opcode===0xd7);
  expect(indices.length).toBe(2);expect(dbs.length).toBe(2);
  const bytes=source.slice();new DataView(bytes.buffer).setUint32(indices[0]!.dataOffset+16,dbs[1]!.headerOffset,true);
  const result=editXlsBlankWorkbookStream(bytes,{row:0,col:0,value:5});expect(result).toMatchObject({status:'unchanged',reason:'malformed-records'});expect(result.bytes).toBe(bytes);
 });
});

describe('strict raw MULBLANK sibling preservation',()=>{
 const bytes=new Uint8Array([0xbe,0,14,0,7,0,10,0,15,0,16,0,17,0,18,0,13,0]);
 it.each([10,11,12,13])('preserves exact neighboring XF bytes for column %i',col=>{
  const result=replaceXlsBlankRecord(bytes,readRecords(bytes,0,bytes.length)[0]!,col,1.25), view=new DataView(result.bytes.buffer), remaining=new Map<number,number>();
  for(const r of readRecords(result.bytes,0,result.bytes.length)){
   const first=view.getUint16(r.dataOffset+2,true);
   if(r.opcode===0x203){expect(first).toBe(col);expect(view.getUint16(r.dataOffset+4,true)).toBe(15+col-10);expect(view.getFloat64(r.dataOffset+6,true)).toBe(1.25);continue;}
   const count=r.opcode===0xbe?(r.length-6)/2:1;
   for(let i=0;i<count;i++) remaining.set(first+i,view.getUint16(r.dataOffset+4+i*2,true));
  }
  expect(remaining.size).toBe(3);for(let i=10;i<=13;i++)if(i!==col)expect(remaining.get(i)).toBe(15+i-10);
 });
 it('rejects wrong opcode/framing/ranges/nonfinite values before writing',()=>{
  const target=readRecords(bytes,0,bytes.length)[0]!;
  for(const col of [-1,9,14,256,NaN])expect(()=>replaceXlsBlankRecord(bytes,target,col,1)).toThrow();
  for(const value of [NaN,Infinity,-Infinity])expect(()=>replaceXlsBlankRecord(bytes,target,10,value)).toThrow();
  for(const t of [{...target,opcode:0x203},{...target,length:target.length+2},{...target,dataOffset:0},{...target,headerOffset:-1}])expect(()=>replaceXlsBlankRecord(bytes,t,10,1)).toThrow();
 });
});

/** Authored physical pointer fixture: 32 rows then a final single-row DB block. */
function twoBlocks(): Uint8Array {
 const record=(opcode:number,data:Uint8Array)=>{const b=new Uint8Array(data.length+4),v=new DataView(b.buffer);v.setUint16(0,opcode,true);v.setUint16(2,data.length,true);b.set(data,4);return b;};
 const bof=(kind:number)=>{const b=new Uint8Array(16),v=new DataView(b.buffer);v.setUint16(0,0x600,true);v.setUint16(2,kind,true);return record(0x809,b);};
 const bound=record(0x85,new Uint8Array([0,0,0,0,0,0,1,0,83])), index=record(0x20b,new Uint8Array(24));
 const parts=[bof(5),bound,record(0xfc,new Uint8Array(8)),record(0xa,new Uint8Array()),bof(0x10),index,record(0x55,new Uint8Array([8,0]))];
 for(const [first,count]of [[0,32],[32,1]]){
  const blockStart=parts.reduce((n,b)=>n+b.length,0);
  for(let i=0;i<count!;i++){const b=new Uint8Array(16),v=new DataView(b.buffer);v.setUint16(0,first!+i,true);v.setUint16(4,1,true);parts.push(record(0x208,b));}
  for(let i=0;i<count!;i++){const b=new Uint8Array(6),v=new DataView(b.buffer);v.setUint16(0,first!+i,true);v.setUint16(4,15,true);parts.push(record(0x201,b));}
  const db=new Uint8Array(4+2*count!),v=new DataView(db.buffer);v.setUint32(0,parts.reduce((n,b)=>n+b.length,0)-blockStart,true);
  for(let i=0;i<count!;i++)v.setUint16(4+2*i,i===0?(count!-1)*20:10,true);
  parts.push(record(0xd7,db));
 }
 parts.push(record(0xa,new Uint8Array()));
 const bytes=new Uint8Array(parts.reduce((n,b)=>n+b.length,0));let offset=0;for(const p of parts){bytes.set(p,offset);offset+=p.length;}
 const records=readRecords(bytes,0,bytes.length),v=new DataView(bytes.buffer),b=records.find(r=>r.opcode===0x85)!,ix=records.find(r=>r.opcode===0x20b)!;
 v.setUint32(b.dataOffset,records.filter(r=>r.opcode===0x809)[1]!.headerOffset,true);
 v.setUint32(ix.dataOffset+8,33,true);v.setUint32(ix.dataOffset+12,records.find(r=>r.opcode===0x55)!.headerOffset,true);
 records.filter(r=>r.opcode===0xd7).forEach((r,i)=>v.setUint32(ix.dataOffset+16+i*4,r.headerOffset,true));
 return bytes;
}
describe('multiple physical row blocks',()=>{
 it.each([0,31,32])('relocates INDEX/DBCELL for blank row %i with other row blocks intact',row=>{
  const input=twoBlocks(),result=editXlsBlankWorkbookStream(input,{row,col:0,value:19.5});expect(result.status).toBe('edited');
  const records=readRecords(result.bytes,0,result.bytes.length),v=new DataView(result.bytes.buffer),byOffset=new Map(records.map(r=>[r.headerOffset,r]));
  const index=records.find(r=>r.opcode===0x20b)!;
  for(let p=index.dataOffset+16;p<index.dataOffset+index.length;p+=4)expect(byOffset.get(v.getUint32(p,true))?.opcode).toBe(0xd7);
  for(const db of records.filter(r=>r.opcode===0xd7)){
   const first=byOffset.get(db.headerOffset-v.getUint32(db.dataOffset,true))!;expect(first.opcode).toBe(0x208);let base=first.dataOffset+first.length;
   for(let i=0;i<(db.length-4)/2;i++){const address=base+v.getUint16(db.dataOffset+4+i*2,true),cell=byOffset.get(address)!;expect(cell).toBeDefined();expect(v.getUint16(cell.dataOffset,true)).toBe(v.getUint16(first.dataOffset,true)+i);base=address;}
  }
 });
 it('refuses duplicate INDEX DB block references and a row without an addressable first cell',()=>{
  const source=twoBlocks(),records=readRecords(source,0,source.length),index=records.find(r=>r.opcode===0x20b)!,db=records.find(r=>r.opcode===0xd7)!;
  for(const mutate of [(v:DataView)=>v.setUint32(index.dataOffset+20,v.getUint32(index.dataOffset+16,true),true),(v:DataView)=>v.setUint16(db.dataOffset+4,0,true)]){
   const bytes=source.slice();mutate(new DataView(bytes.buffer));const result=editXlsBlankWorkbookStream(bytes,{row:0,col:0,value:1});expect(result).toMatchObject({status:'unchanged',reason:'malformed-records'});expect(result.bytes).toBe(bytes);
  }
 });
});
