/** Existing-cell BIFF8 string edits without rebuilding worksheet cell tables.
 * All original SST/CONTINUE bytes, formulas, ROW/DBCELL and unknown cell values
 * survive. Known absolute Workbook pointers are relocated; unknown/future
 * pointer-bearing records are refused before resizing.
 */
import { type BiffRecord } from './legacy-excel-biff8.js';
import { appendXlsSstString, readPreservedXlsSst } from './legacy-excel-preserved-sst.js';
import { KNOWN_XLS_RECORDS, XLS_FRT_WRAPPER } from './legacy-excel-known-records.js';

export type XlsPreservedStringFailure = 'invalid-edit' | 'unsupported-workbook' | 'malformed-records' | 'sheet-not-found' | 'cell-not-supported' | 'ambiguous-cell' | 'unsupported-pointer-record' | 'container-not-writable';
export type XlsPreservedStringResult = {status:'edited';bytes:Uint8Array;recalculationRequired:true} | {status:'unchanged';bytes:Uint8Array;reason:XlsPreservedStringFailure};
type PhysicalRecord = BiffRecord & {depth:number};
type Model = {records:PhysicalRecord[]; worksheets:Array<{start:number;end:number}>; globalsEnd:number};
class EditError extends Error { constructor(readonly reason:XlsPreservedStringFailure) { super(reason); } }
function reject(reason:XlsPreservedStringFailure):never { throw new EditError(reason); }

function model(bytes:Uint8Array):Model {
 const view = new DataView(bytes.buffer,bytes.byteOffset,bytes.byteLength);
 const records:PhysicalRecord[]=[];
 const streams=new Map<number,{start:number;end:number;kind:number}>();
 let offset=0,depth=0,globalsEnd=-1;
 let top:{start:number;kind:number}|undefined;
 while(offset<bytes.length) {
  if(depth===0 && globalsEnd>=0 && bytes[offset]===0 && bytes.subarray(offset).every(b=>b===0)) break;
  if(offset+4>bytes.length) reject('malformed-records');
  const opcode=view.getUint16(offset,true),length=view.getUint16(offset+2,true),dataOffset=offset+4;
  if(length>8224 || dataOffset+length>bytes.length) reject('malformed-records');
  if(opcode===0x0809) {
   if(length<4 || view.getUint16(dataOffset,true)!==0x0600 || depth>=64) reject('unsupported-workbook');
   const kind=view.getUint16(dataOffset+2,true);
   if(depth===0) {
    if(offset===0 ? kind!==5 : globalsEnd<0 || kind===5) reject('malformed-records');
    top={start:offset,kind};
   }
   depth++;
  } else if(opcode===0x000a) {
   if(depth===0 || length!==0) reject('malformed-records');
  } else if(depth===0) reject('malformed-records');
  records.push({opcode,length,headerOffset:offset,dataOffset,depth});
  if(opcode===0x000a && --depth===0) {
   if(!top) reject('malformed-records');
   streams.set(top.start,{...top,end:dataOffset});
   if(top.kind===5) globalsEnd=dataOffset;
   top=undefined;
  }
  offset=dataOffset+length;
 }
 if(depth!==0 || globalsEnd<0) reject('malformed-records');
 if(records.some(r=>r.opcode===0x002f && r.headerOffset<globalsEnd)) reject('unsupported-workbook');
 const bounds=records.filter(r=>r.opcode===0x0085 && r.headerOffset<globalsEnd && r.depth===1);
 if(bounds.length===0) reject('malformed-records');
 const worksheets:Array<{start:number;end:number}>=[];
 const used=new Set<number>();
 for(const bound of bounds) {
  const p=bound.dataOffset;
  if(bound.length<8) reject('malformed-records');
  const chars=view.getUint8(p+6),wide=Boolean(view.getUint8(p+7)&1),kind=view.getUint8(p+5);
  if(chars<1 || chars>31 || bound.length!==8+chars*(wide?2:1) || (view.getUint8(p+4)&3)>2) reject('malformed-records');
  const expected:Record<number,number>={0:0x10,1:0x40,2:0x20,6:6};
  const stream=streams.get(view.getUint32(p,true));
  if(!stream || stream.kind!==expected[kind] || used.has(stream.start)) reject('malformed-records');
  used.add(stream.start);
  if(kind===0) worksheets.push(stream);
 }
 if(used.size!==streams.size-1) reject('malformed-records');
 return {records,worksheets,globalsEnd};
}

function findTarget(bytes:Uint8Array,m:Model,edit:{row:number;col:number;worksheetIndex?:number}):PhysicalRecord {
 const range=m.worksheets[edit.worksheetIndex??0];
 if(!range) reject('sheet-not-found');
 const view=new DataView(bytes.buffer,bytes.byteOffset,bytes.byteLength);
 const targets:PhysicalRecord[]=[];
 for(const r of m.records) {
  if(r.headerOffset<range.start || r.headerOffset>=range.end || r.depth!==1) continue;
  const p=r.dataOffset;
  if([0x0203,0x027e,0x00fd,0x0006,0x0201,0x0204,0x0205,0x00d6].includes(r.opcode)) {
   if(r.length<6) reject('malformed-records');
   if(view.getUint16(p,true)===edit.row && view.getUint16(p+2,true)===edit.col) targets.push(r);
  } else if(r.opcode===0x00bd || r.opcode===0x00be) {
   const stride=r.opcode===0x00bd?6:2;
   if(r.length<6+stride || (r.length-6)%stride!==0) reject('malformed-records');
   const first=view.getUint16(p+2,true),last=view.getUint16(p+r.length-2,true);
   if(last>255 || last<first || last-first+1!==(r.length-6)/stride) reject('malformed-records');
   if(view.getUint16(p,true)===edit.row && edit.col>=first && edit.col<=last) targets.push(r);
  }
 }
 if(targets.length>1) reject('ambiguous-cell');
 const target=targets[0];
 if(!target || ![0x00fd,0x027e].includes(target.opcode)) reject('cell-not-supported');
 if(target.length!==10) reject('malformed-records');
 return target;
}

type Splice={start:number;end:number;bytes:Uint8Array};
function extSst(entries:ReadonlyArray<{offset:number;recordOffset:number}>,map:(offset:number)=>number):Uint8Array {
 const interval=Math.max(Math.floor(entries.length/128)+1,8);
 const count=Math.ceil(entries.length/interval);
 const out=new Uint8Array(6+count*8),view=new DataView(out.buffer);
 view.setUint16(0,0x00ff,true);view.setUint16(2,out.length-4,true);view.setUint16(4,interval,true);
 for(let i=0;i<count;i++) {
  const entry=entries[i*interval]!;
  const offset=map(entry.offset),base=map(entry.recordOffset);
  if(offset<base || offset-base>65535) reject('malformed-records');
  view.setUint32(6+i*8,offset,true);view.setUint16(10+i*8,offset-base,true);
 }
 return out;
}

/** Edit an existing LABELSST/RK cell in a bare BIFF8 Workbook stream. Worksheet
 * indices follow tab order, excluding charts/macros. Original rich strings are
 * retained; plain replacement strings are appended or reuse an existing plain
 * entry. No formulas are executed or cell-table records rebuilt.
 */
export function editXlsStringWorkbookStream(input:Uint8Array,edit:{row:number;col:number;value:string;worksheetIndex?:number}):XlsPreservedStringResult {
 const unchanged=(reason:XlsPreservedStringFailure):XlsPreservedStringResult=>({status:'unchanged',bytes:input,reason});
 if(!Number.isInteger(edit.row)||edit.row<0||edit.row>65535||!Number.isInteger(edit.col)||edit.col<0||edit.col>255||!Number.isInteger(edit.worksheetIndex??0)||(edit.worksheetIndex??0)<0||edit.value.length>32767||/[\uD800-\uDBFF](?![\uDC00-\uDFFF])|(?:^|[^\uD800-\uDBFF])[\uDC00-\uDFFF]/u.test(edit.value)) return unchanged('invalid-edit');
 try {
  const m=model(input),target=findTarget(input,m,edit);
  const sst=readPreservedXlsSst(input);
  if(m.records.filter(r=>r.opcode===0x00fc).length!==1 || m.records.find(r=>r.headerOffset===sst.start)?.depth!==1) reject('malformed-records');
  const sourceView=new DataView(input.buffer,input.byteOffset,input.byteLength);
  if(sst.total>0x7fffffff || sst.entries.length>0x7fffffff || m.records.filter(r=>r.opcode===0x00fd).length!==sst.total) reject('malformed-records');
  if(m.records.some(r=>r.opcode===0x00fd && (r.length!==10 || sourceView.getUint32(r.dataOffset+6,true)>=sst.entries.length))) reject('malformed-records');
  if(target.opcode===0x027e && sst.total===0x7fffffff) reject('malformed-records');
  let index=sst.entries.findIndex(e=>e.text===edit.value && e.richRuns===0 && e.extendedBytes===0);
  const splices:Splice[]=[];
  let ext:PhysicalRecord|undefined;
  if(index<0) {
   if(m.records.some(r=>!KNOWN_XLS_RECORDS.has(r.opcode)||r.opcode===XLS_FRT_WRAPPER)) reject('unsupported-pointer-record');
   index=sst.entries.length;
   if(index===0x7fffffff) reject('malformed-records');
   const appended=appendXlsSstString(edit.value);
   splices.push({start:sst.end,end:sst.end,bytes:appended});
   const exts=m.records.filter(r=>r.opcode===0x00ff);
   if(exts.length>1) reject('malformed-records');
   ext=exts[0];
   if(ext) {
    const interval=sourceView.getUint16(ext.dataOffset,true);
    if(interval!==Math.max(Math.floor(sst.entries.length/128)+1,8) || ext.length!==2+Math.ceil(sst.entries.length/interval)*8) reject('malformed-records');
    for(let i=0;i<(ext.length-2)/8;i++) {
     const entry=sst.entries[i*interval]!;
     if(sourceView.getUint32(ext.dataOffset+2+i*8,true)!==entry.offset || sourceView.getUint16(ext.dataOffset+6+i*8,true)!==entry.offset-entry.recordOffset) reject('malformed-records');
    }
    if(ext.headerOffset>=m.globalsEnd || ext.depth!==1 || ext.length<2 || (ext.length-2)%8!==0 || m.records[m.records.indexOf(ext)+1]?.opcode===0x003c) reject('malformed-records');
    const placeholder=extSst([...sst.entries,{offset:sst.end+4,recordOffset:sst.end}],x=>x);
    splices.push({start:ext.headerOffset,end:ext.dataOffset+ext.length,bytes:placeholder});
   }
  }
  splices.sort((a,b)=>a.start-b.start||a.end-b.end);
  const map=(offset:number):number=>{
   let moved=offset;
   for(const splice of splices) {
    if(offset>splice.start && offset<splice.end) reject('malformed-records');
    if(offset>=splice.end) moved+=splice.bytes.length-(splice.end-splice.start);
   }
   if(moved<0 || moved>0xffffffff) reject('malformed-records');
   return moved;
  };
  if(ext) {
   // The appended entry lies inside the inserted bytes, so translate its new
   // absolute position separately rather than mapping it as an original byte.
   const entries=sst.entries.map(e=>({offset:map(e.offset),recordOffset:map(e.recordOffset)}));
   const insertionBase=sst.end+splices.filter(s=>s.end<=sst.end && s.start!==sst.end).reduce((n,s)=>n+s.bytes.length-(s.end-s.start),0);
   entries.push({offset:insertionBase+4,recordOffset:insertionBase});
   splices.find(s=>s.start===ext!.headerOffset && s.end> s.start)!.bytes=extSst(entries,x=>x);
  }
  const totalLength=input.length+splices.reduce((n,s)=>n+s.bytes.length-(s.end-s.start),0);
  const out=new Uint8Array(totalLength);
  let from=0,to=0;
  for(const splice of splices) {
   out.set(input.subarray(from,splice.start),to);to+=splice.start-from;
   out.set(splice.bytes,to);to+=splice.bytes.length;from=splice.end;
  }
  out.set(input.subarray(from),to);
  const view=new DataView(out.buffer);
  view.setUint16(map(target.headerOffset),0x00fd,true);
  view.setUint32(map(target.dataOffset)+6,index,true);
  view.setUint32(map(sst.start)+4,sst.total+(target.opcode===0x027e?1:0),true);
  view.setUint32(map(sst.start)+8,sst.entries.length+(index===sst.entries.length?1:0),true);
  if(splices.length) {
   const byOffset=new Map(m.records.map(r=>[r.headerOffset,r]));
   for(const record of m.records) {
    if(record.opcode===0x0085 && (record.headerOffset>=m.globalsEnd || record.depth!==1)) reject('malformed-records');
    if(record.opcode===0x0085) view.setUint32(map(record.dataOffset),map(sourceView.getUint32(record.dataOffset,true)),true);
    if(record.opcode===0x020b) {
     if(record.length<16 || (record.length-16)%4!==0) reject('malformed-records');
     for(let p=record.dataOffset+12;p<record.dataOffset+record.length;p+=4) {
      const address=sourceView.getUint32(p,true);
      const pointed=byOffset.get(address);
      if(!pointed || pointed.opcode!==(p===record.dataOffset+12?0x0055:0x00d7)) reject('malformed-records');
      view.setUint32(map(p),map(address),true);
     }
    }
   }
  }
  return {status:'edited',bytes:out,recalculationRequired:true};
 } catch(error) { return unchanged(error instanceof EditError?error.reason:'malformed-records'); }
}
