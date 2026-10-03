/** Shared strict BIFF8 physical framing, tab selection and edit guards. */
import type { BiffRecord } from './legacy-excel-biff8.js';
import type { XlsPreservedStringFailure } from './legacy-excel-preserved-string-edit.js';
export type PhysicalRecord = BiffRecord & {depth:number};
export type Model = {records:PhysicalRecord[]; worksheets:Array<{start:number;end:number}>; globalsEnd:number};
export class EditError extends Error { constructor(readonly reason:XlsPreservedStringFailure) { super(reason); } }
export function reject(reason:XlsPreservedStringFailure):never { throw new EditError(reason); }

export function model(bytes:Uint8Array):Model {
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

export function findTarget(bytes:Uint8Array,m:Model,edit:{row:number;col:number;worksheetIndex?:number},allowed:readonly number[]=[0x00fd,0x027e,0x0203,0x00bd]):PhysicalRecord {
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
   if(r.opcode===0x00be && (first>254 || last<=first)) reject('malformed-records');
   if(view.getUint16(p,true)===edit.row && edit.col>=first && edit.col<=last) targets.push(r);
  }
 }
 if(targets.length>1) reject('ambiguous-cell');
 const target=targets[0];
 if(!target || !allowed.includes(target.opcode)) reject('cell-not-supported');
 if(target.opcode!==0x00bd && target.opcode!==0x00be && target.length!==(target.opcode===0x0203?14:target.opcode===0x0205?8:target.opcode===0x0201?6:10)) reject('malformed-records');
 for(const record of m.records.filter(r=>r.headerOffset>=range.start&&r.headerOffset<range.end&&r.depth===1&&[0x0221,0x04bc,0x0236].includes(r.opcode))) {
  if(record.length<6) reject('malformed-records');
  const p=record.dataOffset;
  if(edit.row>=view.getUint16(p,true)&&edit.row<=view.getUint16(p+2,true)&&edit.col>=view.getUint8(p+4)&&edit.col<=view.getUint8(p+5)) reject('cell-not-supported');
 }
 return target;
}
