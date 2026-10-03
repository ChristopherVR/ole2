/** Binary VSD v11 layout facts: Apache POI HDGF and LibreOffice libvisio.
 * https://github.com/apache/poi/tree/trunk/poi-scratchpad/src/main/java/org/apache/poi/hdgf
 * https://github.com/LibreOffice/libvisio/blob/master/src/lib/VSDParser.cpp
 * These are binary implementation references, not the XML MS-VSDX spec.
 * This reader exposes explicit stored values, without master/style evaluation. */
import { readCompoundFileStream } from './ole2-stream-edit.js';
import { decodeVsdBlock } from './vsd-compression.js';

export class VsdError extends Error {
 constructor(readonly reason: string) { super(`Unsupported or invalid VSD: ${reason}`); this.name = 'VsdError'; }
}
export interface VsdTransform { pinX: number; pinY: number; width: number; height: number; localPinX: number; localPinY: number; angle: number; flipX: boolean; flipY: boolean }
export interface VsdGeometry { readonly kind: 'moveTo' | 'lineTo'; readonly id: number; readonly x: number; readonly y: number }
export interface VsdRecordLocation { block: VsdBlock; offset: number; length: number }
export interface VsdShapeData { id: number; kind: 'shape' | 'group' | 'foreign'; coordinateSpace: 'shape-local'; parentId?: number; masterPageId?: number; masterShapeId?: number; unsupportedGeometry: boolean; unsafeTransform: boolean; text?: string; transform?: VsdTransform; geometry: VsdGeometry[]; textRecord?: VsdRecordLocation; transformRecord?: VsdRecordLocation; unsafeText: boolean }
export interface VsdPageData { id: number; background: boolean; width?: number; height?: number; scale?: number; shapes: VsdShapeData[] }
export interface VsdBlock { offset: number; length: number; format: number; type: number; bytes: Uint8Array; parent?: VsdBlock; pointerOffset: number }
export interface VsdDrawingData { stream: Uint8Array; streamPath: string[]; version: 11; pages: VsdPageData[]; blocks: VsdBlock[]; writable: boolean }
export const VSD_MAX_BYTES = 64 * 1024 * 1024;
const MAX_BYTES = VSD_MAX_BYTES;
export function validateVsdInput(input: Uint8Array): Uint8Array {
 if (!(input instanceof Uint8Array) || input.byteLength > MAX_BYTES) throw new VsdError('input-budget'); return input;
}
function view(bytes: Uint8Array): DataView { return new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength); }
function bounded(bytes: Uint8Array, offset: number, length: number): void {
 if (!Number.isSafeInteger(offset) || offset < 0 || length < 0 || offset > bytes.length || length > bytes.length - offset) throw new VsdError('record-bounds');
}
function numeric(bytes: Uint8Array, offset: number): number {
 bounded(bytes, offset, 8); const value = view(bytes).getFloat64(offset, true);
 if (!Number.isFinite(value)) throw new VsdError('non-finite-geometry'); return value;
}
export function readVsdDrawing(input: Uint8Array): VsdDrawingData {
 validateVsdInput(input);
 for (const name of ['EncryptedPackage', 'EncryptionInfo']) if (readCompoundFileStream(input, [name])) throw new VsdError('encrypted-container');
 const streamPath = ['VisioDocument'], candidateStream = readCompoundFileStream(input, streamPath);
 if (!candidateStream || candidateStream.length < 54 || candidateStream.length > MAX_BYTES) throw new VsdError('missing-or-oversized-visio-stream');
 const stream = candidateStream;
 if (new TextDecoder().decode(stream.subarray(0, 20)) !== 'Visio (TM) Drawing\r\n') throw new VsdError('signature');
 const header = view(stream), version = header.getUint16(26, true);
 if (version !== 11) throw new VsdError(`version-${version}`);
 const size = header.getUint32(28, true);
 if (size < 54 || size > stream.length) throw new VsdError('document-size');
 const data: VsdDrawingData = { stream, streamPath, version: 11, pages: [], blocks: [], writable: true };
 const active = new Set<number>(), seen = new Set<number>(); let decodedBudget = MAX_BYTES, records = 0;
 function visit(ownerBytes: Uint8Array, pointerOffset: number, parent?: VsdBlock, page?: VsdPageData, shape?: VsdShapeData, depth = 0, id = 0): void {
  if (depth > 64 || ++records > 100000) throw new VsdError('traversal-limit');
  bounded(ownerBytes, pointerOffset, 18); const pv = view(ownerBytes);
  const type = pv.getUint32(pointerOffset, true); if (!type) return;
  if(type===0x4d) throw new VsdError('unsupported-guide');
  if(shape && [0x46,0x4a].includes(type)) throw new VsdError('unsupported-sheet-owner');
  const offset = pv.getUint32(pointerOffset + 8, true), length = pv.getUint32(pointerOffset + 12, true), format = pv.getUint16(pointerOffset + 16, true);
  // Reserved empty pointers occur in real drawings (e.g. type 0xff/format 0x60).
  if (offset === 0 && length === 0) return;
  if (offset < 54 || !length || offset > size || length > size - offset) throw new VsdError('pointer-bounds');
  if (active.has(offset)) throw new VsdError('pointer-cycle');
  if (seen.has(offset)) { data.writable = false; return; } seen.add(offset); active.add(offset);
  const compressed = (format & 2) !== 0, shift = compressed ? 4 : 0;
  const bytes = compressed ? decodeVsdBlock(stream.subarray(offset, offset + length), decodedBudget) : stream.slice(offset, offset + length);
  decodedBudget -= bytes.length; if (decodedBudget < 0) throw new VsdError('decoded-budget');
  const block: VsdBlock = { offset, length, format, type, bytes, parent, pointerOffset }; data.blocks.push(block);
  if (type === 0x15) { page = { id, background: !(format & 1), shapes: [] }; data.pages.push(page); shape = undefined; }
  if (page && [0x47, 0x48, 0x4e].includes(type)) {
   shape = { id, kind: type === 0x47 ? 'group' : type === 0x4e ? 'foreign' : 'shape', coordinateSpace:'shape-local', unsupportedGeometry: false, unsafeTransform: false, geometry: [], unsafeText: false }; page.shapes.push(shape);
   if (bytes.length - shift >= 54) { const sv=view(bytes);shape.parentId=sv.getUint32(shift+10,true);shape.masterPageId=sv.getUint32(shift+18,true);shape.masterShapeId=sv.getUint32(shift+26,true); }
  }
  const mode = format >>> 4;
  if (mode === 5 || type === 0x14) {
   bounded(bytes, shift, 4); const table = view(bytes).getUint32(shift, true) + shift - 4;
   bounded(bytes, table, 12); const count = view(bytes).getInt32(table + 4, true), order = view(bytes).getUint32(table, true);
   if (count < 0 || count > 100000 || count > Math.floor((bytes.length - table - 12) / 18)) throw new VsdError('pointer-count');
   const orderLength = order <= 1 ? 0 : order;
   bounded(bytes, table + 12 + count * 18, orderLength * 4);
   const indexes = Array.from({ length: count }, (_, i) => i);
   if (orderLength) {
    const ranks = new Map<number, number>();
    for (let i = 0; i < orderLength; i++) { const idx = view(bytes).getUint32(table + 12 + count * 18 + i * 4, true); if (idx >= count || ranks.has(idx)) throw new VsdError('pointer-order'); ranks.set(idx, i); }
    indexes.sort((a,b) => (ranks.get(a) ?? count + a) - (ranks.get(b) ?? count + b));
   }
   for (const idx of indexes) visit(bytes, table + 12 + idx * 18, block, page, shape, depth + 1, idx);
  } else if (page && [8, 12, 13].includes(mode)) {
   const stack: { level: number; shape: VsdShapeData }[] = [];
   // Compressed chunk streams begin directly with their first chunk header;
   // only compressed blob/pointer-table streams have the four-byte prefix.
   let cursor = 0;
   while (cursor < bytes.length) {
    while (cursor < bytes.length && bytes[cursor] === 0) cursor++;
    if (cursor === bytes.length) break;
    if (++records > 100000) throw new VsdError('record-count');
    bounded(bytes, cursor, 19); const cv = view(bytes), chunk = cv.getUint32(cursor, true), chunkId = cv.getUint32(cursor + 4, true), list = cv.getUint32(cursor + 8, true), length = cv.getUint32(cursor + 12, true), level = cv.getUint16(cursor + 16, true), marker = bytes[cursor + 18]!;
    if(chunk===0x4d) throw new VsdError('unsupported-guide');
    let trailer = list || [0x71,0x70,0x6b,0x6a,0x69,0x66,0x65,0x2c].includes(chunk) ? 8 : 0;
    if (list || (level === 2 && marker === 0x55) || (level === 2 && marker === 0x54 && chunk === 0xaa) || (level === 3 && marker !== 0x50 && marker !== 0x54)) trailer += 4;
    if ([0x64,0x65,0x66,0x69,0x6a,0x6b,0x6f,0x71,0x92,0xa9,0xb4,0xb6,0xb9,0xc7].includes(chunk) && trailer !== 12 && trailer !== 4) trailer += 4;
    if ([0x1f,0xc9,0x2d,0xd1].includes(chunk)) trailer = 0;
    const start = cursor + 19;
    if(length + trailer > bytes.length - start) throw new VsdError(`chunk-bounds-${offset}-${cursor}-${chunk}-${length}-${trailer}-${bytes.length}`);
    while (stack.length && level <= stack[stack.length - 1]!.level) stack.pop();
    let current = stack[stack.length - 1]?.shape ?? shape;
    if(current && [0x46,0x4a].includes(chunk)) throw new VsdError('unsupported-sheet-owner');
    if (page && [0x47,0x48,0x4e].includes(chunk)) {
     current = { id: chunkId, kind: chunk === 0x47 ? 'group' : chunk === 0x4e ? 'foreign' : 'shape', coordinateSpace:'shape-local', unsupportedGeometry: false, unsafeTransform: false, geometry: [], unsafeText: false };
     if(length<54) throw new VsdError('shape-header');
     current.parentId=cv.getUint32(start+10,true);current.masterPageId=cv.getUint32(start+18,true);current.masterShapeId=cv.getUint32(start+26,true);
     page.shapes.push(current); stack.push({ level, shape: current });
    }
    if (chunk === 0x92 && page) {
     if(length<54) throw new VsdError('page-properties');
     page.width=numeric(bytes,start+1);page.height=numeric(bytes,start+10);
     const pageScale=numeric(bytes,start+37),drawingScale=numeric(bytes,start+46);
     if(page.width<0 || page.height<0 || pageScale<=0 || drawingScale<=0) throw new VsdError('invalid-page-scale');
     page.scale=pageScale/drawingScale;if(!Number.isFinite(page.scale))throw new VsdError('invalid-page-scale');
    }
    if(current && chunk===0x89 && length>0 && bytes[start]!==0)current.unsupportedGeometry=true;
    if (current && chunk === 0x0e) {
     if (length < 8 || (length - 8) % 2 || current.textRecord) throw new VsdError('text-record');
     current.text = new TextDecoder('utf-16le', { fatal: true, ignoreBOM: true }).decode(bytes.subarray(start + 8, start + length));
     current.textRecord = { block, offset: start, length };
    }
    if (current && chunk === 0x9b) {
     if (length < 65 || current.transformRecord) throw new VsdError('transform-record');
     const values = Array.from({length: 7}, (_, i) => numeric(bytes, start + 1 + i * 9));
     current.transform = {pinX: values[0]!,pinY:values[1]!,width:values[2]!,height:values[3]!,localPinX:values[4]!,localPinY:values[5]!,angle:values[6]!,flipX:!!bytes[start+63],flipY:!!bytes[start+64]};
     current.transformRecord = { block, offset: start, length };
    }
    if (current && [0x8a,0x8b].includes(chunk)) { if (length < 18) throw new VsdError('geometry-record'); current.geometry.push({kind:chunk===0x8a?'moveTo':'lineTo',id:chunkId,x:numeric(bytes,start+1),y:numeric(bytes,start+10)}); }
    if (current && [0xa1,0x6f].includes(chunk)) current.unsafeText = true;
    if (current && [0x8c,0x8d,0x8f,0x90,0xa5,0xa6,0xc1,0xc3].includes(chunk)) current.unsupportedGeometry=true;
    if (current && (chunk===0x9d || chunk===0xa4)) current.unsafeTransform=true;
    cursor = start + length + trailer;
   }
  }
  active.delete(offset);
 }
 if (header.getUint32(36, true) !== 0x14) throw new VsdError('trailer-type');
 visit(stream, 36);
 const ranges = [...data.blocks].sort((a,b)=>a.offset-b.offset);
 for (let i=1;i<ranges.length;i++) if (ranges[i]!.offset < ranges[i-1]!.offset+ranges[i-1]!.length) data.writable=false;
 if (!data.pages.length) throw new VsdError('no-decoded-pages');
 const pageIds = new Set<number>();
 for (const page of data.pages) { if(pageIds.has(page.id)) throw new VsdError('duplicate-page-id'); pageIds.add(page.id); const ids = new Set<number>(); for (const shape of page.shapes) { if (ids.has(shape.id)) throw new VsdError('duplicate-shape-id'); ids.add(shape.id); } }
 return data;
}
