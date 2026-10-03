import { buildOle2 } from '../../../src/ole2-parser-write.js';
import { encodeVsdBlock } from '../../../src/vsd-compression.js';
/** Authored here from independently researched binary record layout facts.
 * No third-party fixture bytes. These validate format mechanics; they do not
 * establish native Visio acceptance. Oracle corpus tests are separate. */
export function makeVsdFixture(options: { compressed?: boolean; mutate?: (stream: Uint8Array)=>void } = {}): Uint8Array {
 const compressed = options.compressed ?? true, mutate=options.mutate;
 const parts: Uint8Array[] = [], offsets: number[]=[]; let size=54;
 function add(bytes:Uint8Array, compression=compressed) {const stored=compression?encodeVsdBlock(bytes):bytes;offsets.push(size);parts.push(stored);size+=stored.length;return {offset:offsets.at(-1)!,length:stored.length};}
 function chunk(type:number,id:number,payload:Uint8Array,level=2,trailer=0) {const bytes=new Uint8Array(19+payload.length+trailer),v=new DataView(bytes.buffer);v.setUint32(0,type,true);v.setUint32(4,id,true);v.setUint32(12,payload.length,true);v.setUint16(16,level,true);bytes[18]=0x50;bytes.set(payload,19);return bytes;}
 function coords(values:number[]) {const b=new Uint8Array(values.length*9),v=new DataView(b.buffer);values.forEach((n,i)=>v.setFloat64(i*9+1,n,true));return b;}
 const text=new Uint8Array(8+12),tv=new DataView(text.buffer);'Hello\n'.split('').forEach((s,i)=>tv.setUint16(8+i*2,s.charCodeAt(0),true));
 const transform=new Uint8Array(65);transform.set(coords([2,3,4,2,2,1,0]));for(let i=0;i<7;i++)transform[i*9]=i===6?80:64;
 const shapeHeader=new Uint8Array(54),sv=new DataView(shapeHeader.buffer);for(const offset of [10,18,26,34,42,50])sv.setUint32(offset,0xffffffff,true);
 const order=new Uint8Array(12),ov=new DataView(order.buffer);ov.setUint32(4,4,true);ov.setUint32(8,7,true);
 const shapeId=new Uint8Array(4);new DataView(shapeId.buffer).setUint32(0,7,true);
 const chunks=[chunk(0x46,0,new Uint8Array(0),1),chunk(0x92,0,coords([8,11,0,0,1,1]),3,4),chunk(0x65,0,order,2,12),chunk(0x83,7,shapeId,3),chunk(0x48,7,shapeHeader,1),chunk(0x9b,0,transform),chunk(0x0e,0,text),chunk(0x6c,0,new Uint8Array(0)),chunk(0x89,0,Uint8Array.of(0),3),chunk(0x8a,0,coords([0,0]),3),chunk(0x8b,1,coords([4,2]),3),chunk(0xee,3,Uint8Array.of(8,9,7))];
 const page=new Uint8Array(chunks.reduce((n,b)=>n+b.length,0));let cursor=0;for(const c of chunks){page.set(c,cursor);cursor+=c.length;}
 const pageRange=add(page);
 function table(type:number,range:{offset:number,length:number},format:number) {
  const shift=compressed?4:0,b=new Uint8Array(shift+8+12+18),v=new DataView(b.buffer),base=shift+8;
  v.setUint32(shift,12,true);v.setInt32(base+4,1,true);
  const p=base+12;v.setUint32(p,type,true);v.setUint32(p+8,range.offset,true);v.setUint32(p+12,range.length,true);v.setUint16(p+16,format,true);return b;
 }
 const pagesRange=add(table(0x15,pageRange,0xd1+(compressed?2:0))),trailerRange=add(table(0x27,pagesRange,0x50+(compressed?2:0)));
 const stream=new Uint8Array(size),v=new DataView(stream.buffer);stream.set(new TextEncoder().encode('Visio (TM) Drawing\r\n'));v.setUint16(26,11,true);v.setUint32(28,size,true);
 v.setUint32(36,0x14,true);v.setUint32(44,trailerRange.offset,true);v.setUint32(48,trailerRange.length,true);v.setUint16(52,0x50+(compressed?2:0),true);
 parts.forEach((p,i)=>stream.set(p,offsets[i]!));mutate?.(stream);
 return new Uint8Array(buildOle2(new Map([['VisioDocument',stream],['OpaqueUnknown',Uint8Array.of(9,4,8,3,5)]])));
}
