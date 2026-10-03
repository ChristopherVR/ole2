/** Targeted BIFF8 cell-record replacement and validated row-block relocation.
 * No worksheet records are reconstructed outside the selected numeric record.
 */
import type { BiffRecord } from './legacy-excel-biff8.js';

export interface XlsCellRecordReplacement {start:number;end:number;bytes:Uint8Array}
const CELL_OPS=new Set([0x0006,0x0201,0x0203,0x0204,0x0205,0x027e,0x00fd,0x00bd,0x00be,0x00d6]);
function record(opcode:number,data:Uint8Array):Uint8Array {
 const out=new Uint8Array(data.length+4),view=new DataView(out.buffer);
 view.setUint16(0,opcode,true);view.setUint16(2,data.length,true);out.set(data,4);
 return out;
}

/** Preserve every neighboring MULRK cell's exact ixfe/RK bytes. */
export function replaceXlsNumericRecord(input:Uint8Array,target:BiffRecord,col:number,index:number):XlsCellRecordReplacement {
 if(!Number.isInteger(col)||col<0||col>255||!Number.isInteger(index)||index<0||index>0x7fffffff||
    !Number.isSafeInteger(target.headerOffset)||target.headerOffset<0||target.dataOffset!==target.headerOffset+4||
    !Number.isInteger(target.length)||target.length<0||target.dataOffset+target.length>input.length||
    ![0x0203,0x00bd].includes(target.opcode))throw new Error('Invalid numeric record replacement');
 const p=target.dataOffset,view=new DataView(input.buffer,input.byteOffset,input.byteLength);
 if(view.getUint16(target.headerOffset,true)!==target.opcode||view.getUint16(target.headerOffset+2,true)!==target.length)throw new Error('Record framing mismatch');
 if(target.opcode===0x0203 ? target.length!==14||view.getUint16(p+2,true)!==col : target.length<12||(target.length-6)%6!==0)throw new Error('Invalid numeric record size');
 const row=view.getUint16(p,true);
 const first=target.opcode===0x00bd?view.getUint16(p+2,true):col;
 const count=target.opcode===0x00bd?(target.length-6)/6:1;
 const selected=col-first;
 if(target.opcode===0x00bd && (first>255||selected<0||selected>=count||view.getUint16(p+target.length-2,true)!==first+count-1||first+count-1>255))throw new Error('Invalid MULRK column range');
 const data=new Uint8Array(10),cell=new DataView(data.buffer);
 cell.setUint16(0,row,true);cell.setUint16(2,col,true);
 cell.setUint16(4,view.getUint16(p+(target.opcode===0x00bd?4+selected*6:4),true),true);
 cell.setUint32(6,index,true);
 const label=record(0x00fd,data);
 const group=(start:number,size:number):Uint8Array=>{
  if(size===0)return new Uint8Array();
  const bytes=new Uint8Array(size===1?10:6+size*6),dv=new DataView(bytes.buffer);
  dv.setUint16(0,row,true);dv.setUint16(2,first+start,true);
  bytes.set(input.subarray(p+4+start*6,p+4+(start+size)*6),4);
  if(size>1)dv.setUint16(bytes.length-2,first+start+size-1,true);
  return record(size===1?0x027e:0x00bd,bytes);
 };
 const prefix=target.opcode===0x00bd?group(0,selected):new Uint8Array();
 const suffix=target.opcode===0x00bd?group(selected+1,count-selected-1):new Uint8Array();
 const bytes=new Uint8Array(prefix.length+label.length+suffix.length);
 bytes.set(prefix);bytes.set(label,prefix.length);bytes.set(suffix,prefix.length+label.length);
 return{start:target.headerOffset,end:p+target.length,bytes};
}

/** Patch DBCELL only when its row block spans a record-size change. Its ROW and
 * cell addresses are validated against the original records before writing.
 * A first-cell pointer to a split MULRK maps to the first replacement record.
 */
export function relocateXlsDbCells(
 input:Uint8Array,out:Uint8Array,records:readonly BiffRecord[],replacement:XlsCellRecordReplacement,map:(address:number)=>number,
):void {
 const source=new DataView(input.buffer,input.byteOffset,input.byteLength),dest=new DataView(out.buffer,out.byteOffset,out.byteLength);
 const byOffset=new Map(records.map(r=>[r.headerOffset,r]));
 for(const db of records.filter(r=>r.opcode===0x00d7)) {
  if(db.length<4 || (db.length-4)%2!==0 || db.length>68)throw new Error('Invalid DBCELL size');
  const back=source.getUint32(db.dataOffset,true);
  if(back===0) {if(db.length!==4)throw new Error('Unsupported empty DBCELL');continue;}
  const firstRow=byOffset.get(db.headerOffset-back);
  if(!firstRow || firstRow.opcode!==0x0208 || firstRow.length!==16)throw new Error('Invalid DBCELL row pointer');
  if(replacement.start<firstRow.headerOffset || replacement.start>=db.headerOffset)continue;
  const rows=records.filter(r=>r.opcode===0x0208&&r.headerOffset>=firstRow.headerOffset&&r.headerOffset<db.headerOffset);
  if(rows.length!==(db.length-4)/2 || rows.some(r=>r.length!==16))throw new Error('Unsupported DBCELL row block');
  let base=firstRow.dataOffset+firstRow.length;
  const addresses:number[]=[];
  for(let i=0;i<rows.length;i++) {
   const address=base+source.getUint16(db.dataOffset+4+i*2,true),cell=byOffset.get(address);
   if(!cell || !CELL_OPS.has(cell.opcode) || address>=db.headerOffset || source.getUint16(cell.dataOffset,true)!==source.getUint16(rows[i]!.dataOffset,true))throw new Error('Unsupported DBCELL first-cell pointer');
   addresses.push(address);base=address;
  }
  const offset=map(db.headerOffset)-map(firstRow.headerOffset);
  if(offset<0 || offset>0xffffffff)throw new Error('DBCELL row pointer overflow');
  dest.setUint32(map(db.dataOffset),offset,true);
  let newBase=map(firstRow.dataOffset+firstRow.length);
  for(let i=0;i<addresses.length;i++) {
   const address=map(addresses[i]!),relative=address-newBase;
   if(relative<0 || relative>65535)throw new Error('DBCELL cell pointer overflow');
   dest.setUint16(map(db.dataOffset+4+i*2),relative,true);newBase=address;
  }
 }
}
