import { describe, it, expect } from 'vitest';
import { readFileSync } from 'node:fs';
import { buildOle2 } from '../src/ole2-parser-write.js';
import { readVsdDrawing } from '../src/vsd-reader.js';
import { VsdDocument } from '../src/vsd-document.js';

// Entirely authored layout mechanics, not a native fidelity fixture.
function fixture(change?: (chunks: Uint8Array[]) => void, depth = 1): Uint8Array {
 const chunks: Uint8Array[] = [];
 const chunk=(type:number,id:number,payload:Uint8Array,level:number,trailer=0)=>{const b=new Uint8Array(19+payload.length+trailer),v=new DataView(b.buffer);v.setUint32(0,type,true);v.setUint32(4,id,true);v.setUint32(12,payload.length,true);v.setUint16(16,level,true);b[18]=0x50;b.set(payload,19);chunks.push(b);return b;};
 const header=(parent:number)=>{const b=new Uint8Array(54),v=new DataView(b.buffer);for(const o of [10,18,26,34,42,50])v.setUint32(o,o===10?parent:0xffffffff,true);return b;};
 const order=(element:number,actual:number)=>{const b=new Uint8Array(12),v=new DataView(b.buffer);v.setUint32(4,4,true);v.setUint32(8,element,true);chunk(0x65,0,b,2,12);const id=new Uint8Array(4);new DataView(id.buffer).setUint32(0,actual,true);chunk(0x83,element,id,3);};
 chunk(0x46,0,header(0xffffffff),1);const props=new Uint8Array(54),pv=new DataView(props.buffer);[8,11,0,0,1,1].forEach((n,i)=>pv.setFloat64(i*9+1,n,true));chunk(0x92,0,props,3,4);order(123,42);
 chunk(0x48,7,header(41+depth),1);
 for(let i=0;i<depth;i++){chunk(0x47,42+i,header(i===0?0:41+i),1);order(500+i,i===depth-1?7:43+i);}
 change?.(chunks);const page=new Uint8Array(chunks.reduce((n,c)=>n+c.length,0));let at=0;for(const c of chunks){page.set(c,at);at+=c.length;}
 const table=(type:number,offset:number,length:number,format:number)=>{const b=new Uint8Array(38),v=new DataView(b.buffer);v.setUint32(0,12,true);v.setInt32(12,1,true);v.setUint32(20,type,true);v.setUint32(28,offset,true);v.setUint32(32,length,true);v.setUint16(36,format,true);return b;};
 const pages=table(0x15,54,page.length,0xd1),root=table(0x27,54+page.length,38,0x50),stream=new Uint8Array(54+page.length+76),sv=new DataView(stream.buffer);stream.set(new TextEncoder().encode('Visio (TM) Drawing\r\n'));sv.setUint16(26,11,true);sv.setUint32(28,stream.length,true);sv.setUint32(36,0x14,true);sv.setUint32(44,54+page.length+38,true);sv.setUint32(48,38,true);sv.setUint16(52,0x50,true);stream.set(page,54);stream.set(pages,54+page.length);stream.set(root,54+page.length+38);return new Uint8Array(buildOle2(new Map([['VisioDocument',stream],['Unknown',Uint8Array.of(5,8,2)]])));
}
const view=(b:Uint8Array)=>new DataView(b.buffer,b.byteOffset,b.byteLength);
describe('validated VSD hierarchical drawing order',()=>{
 it('resolves element IDs independently of shape IDs and retains physical lookup order',()=>{const bytes=fixture(),doc=new VsdDocument(bytes),page=doc.pages[0]!;expect(page.topLevelShapeIds).toEqual([42]);expect(page.shapes.map(s=>s.id)).toEqual([7,42]);expect(page.shapes[1]!.childShapeIds).toEqual([7]);expect(Object.isFrozen(page.topLevelShapeIds)).toBe(true);expect(page.topLevelShapeIds).not.toBe(page.topLevelShapeIds);expect(doc.serialize()).toEqual(bytes);expect(doc.dirty).toBe(false);});
 it('matches actual native page/group drawing order without flattening groups into roots',()=>{const bytes=new Uint8Array(readFileSync(new URL('./fixtures/vsd/native-hierarchy.vsd',import.meta.url))),doc=new VsdDocument(bytes);expect(doc.pages.map(p=>p.topLevelShapeIds)).toEqual([[2,1,5,6],[1,2,3]]);expect(doc.pages[0]!.shapes.map(s=>s.id)).toEqual([1,2,3,4,5,6]);expect(doc.pages[0]!.shapes.find(s=>s.id===5)!.childShapeIds).toEqual([3,4]);expect(doc.pages[0]!.shapes.find(s=>s.id===6)!.masterPageId).toBe(2);expect(doc.pages[1]!.shapes[0]!.unsupportedGeometry).toBe(true);expect(doc.pages.every(p=>p.shapeOrderIssue===undefined)).toBe(true);expect(doc.serialize()).toEqual(bytes);expect(doc.revision).toBe(0);});
 const mutations: [string,(chunks:Uint8Array[])=>void][]=[
 ['huge child range',c=>view(c[2]!).setUint32(23,0xfffffffc,true)],
 ['misaligned child range',c=>view(c[2]!).setUint32(23,3,true)],
 ['huge subheader',c=>view(c[2]!).setUint32(19,0xffffffff,true)],
 ['separator overrun',c=>view(c[2]!).setUint32(23,16,true)],
 ['mapping owner level',c=>view(c[3]!).setUint16(16,4,true)],
 ['short mapping',c=>view(c[3]!).setUint32(12,3,true)],
 ['missing mapping',c=>view(c[3]!).setUint32(0,0xee,true)],
 ['contradictory parent',c=>view(c[4]!).setUint32(29,999,true)],
 ['duplicate element',c=>{view(c[2]!).setUint32(23,8,true);view(c[2]!).setUint32(31,123,true);}],
 ['duplicate mapping',c=>c.splice(4,0,c[3]!.slice())],
 ['duplicate order',c=>c.splice(4,0,c[2]!.slice(),c[3]!.slice())],
 ];
 it.each(mutations)('reports unresolved %s while preserving every input byte',(_name,mutate)=>{const bytes=fixture(mutate),before=bytes.slice(),page=readVsdDrawing(bytes).pages[0]!;expect(page.topLevelShapeIds).toBeUndefined();expect(page.shapeOrderIssue).toBeTruthy();expect(bytes).toEqual(before);});
 it('bounds iterative hierarchy depth without recursing through shape ownership',()=>{const page=readVsdDrawing(fixture(undefined,66)).pages[0]!;expect(page.topLevelShapeIds).toBeUndefined();expect(page.shapeOrderIssue).toBe('shape-order-depth');});
});

it('rejects aggregate child budgets before building order arrays',()=>{
 const bytes=fixture(chunks=>{const old=chunks[2]!,payloadLength=8+100001*4,b=new Uint8Array(19+payloadLength+12);b.set(old.subarray(0,19));const v=view(b);v.setUint32(12,payloadLength,true);v.setUint32(23,100001*4,true);chunks[2]=b;});
 expect(()=>readVsdDrawing(bytes)).toThrow(/shape-order-budget/);
});
