import { resizeCompoundFileStream } from './ole2-stream-resize.js';
import { encodeVsdBlock } from './vsd-compression.js';
import { readVsdDrawing, VSD_MAX_BYTES, VsdError, type VsdDrawingData, type VsdRecordLocation, type VsdTransform } from './vsd-reader.js';

/** Retain the entire original VisioDocument stream and append only the edited
 * leaf and its rewritten pointer-table ancestors. Unknown records/prefixes and
 * unrelated compound streams remain intact. No flattening or Office execution. */
function replaceRecord(input: Uint8Array, drawing: VsdDrawingData, record: VsdRecordLocation, replacement: Uint8Array): Uint8Array {
 if (!drawing.writable) throw new VsdError('shared-or-overlapping-allocation');
 let block = record.block, decoded = block.bytes.slice(); decoded.set(replacement, record.offset);
 const appended: Uint8Array[] = []; let size = drawing.stream.length;
 const header = drawing.stream.slice();
 for (;;) {
  const stored = block.format & 2 ? encodeVsdBlock(decoded) : decoded;
  if (stored.length > VSD_MAX_BYTES - size) throw new VsdError('writer-size-limit');
  const offset = size; appended.push(stored); size += stored.length;
  const owner = block.parent ? block.parent.bytes.slice() : header;
  const view = new DataView(owner.buffer, owner.byteOffset, owner.byteLength);
  view.setUint32(block.pointerOffset + 8, offset, true); view.setUint32(block.pointerOffset + 12, stored.length, true);
  if (!block.parent) break;
  block = block.parent; decoded = owner;
 }
 const stream = new Uint8Array(size); stream.set(header); let offset = header.length;
 for (const bytes of appended) { stream.set(bytes, offset); offset += bytes.length; }
 new DataView(stream.buffer).setUint32(28, stream.length, true);
 const result = resizeCompoundFileStream(input, drawing.streamPath, stream);
 if (!result.ok) throw new VsdError(`compound-resize-${result.reason}`);
 // Validate the entire reachable graph before the model commits candidate bytes.
 readVsdDrawing(result.bytes); return result.bytes;
}
export function replaceVsdShapeText(input: Uint8Array, pageId: number, shapeId: number, text: string): Uint8Array {
 if (typeof text !== 'string') throw new TypeError('VSD text must be a string');
 const drawing = readVsdDrawing(input), shape = drawing.pages.find(p => p.id === pageId)?.shapes.find(s => s.id === shapeId);
 if (!shape?.textRecord || shape.text === undefined) throw new VsdError('missing-explicit-shape-text');
 if (shape.unsafeText || /[\u0000\ufffc\ufffd]/u.test(shape.text) || /[\u0000\ufffc\ufffd]/u.test(text)) throw new VsdError('text-fields-or-invalid-text');
 if (text.length !== shape.text.length) throw new VsdError('text-length-change');
 for(let i=0;i<text.length;i++) if ((text.charCodeAt(i)<32 || shape.text.charCodeAt(i)<32) && text[i]!==shape.text[i]) throw new VsdError('text-structure-change');
 // Fatal decode also rejects unmatched surrogate input before any allocation.
 const replacement = shape.textRecord.block.bytes.slice(shape.textRecord.offset, shape.textRecord.offset + shape.textRecord.length);
 const view = new DataView(replacement.buffer);
 for (let i = 0; i < text.length; i++) view.setUint16(8 + i * 2, text.charCodeAt(i), true);
 new TextDecoder('utf-16le', {fatal:true}).decode(replacement.subarray(8));
 if (text === shape.text) return input.slice();
 const candidate=replaceRecord(input, drawing, shape.textRecord, replacement);
 if(readVsdDrawing(candidate).pages.find(p=>p.id===pageId)?.shapes.find(s=>s.id===shapeId)?.text!==text) throw new VsdError('text-edit-verification');
 return candidate;
}
export function replaceVsdShapeTransform(input: Uint8Array, pageId: number, shapeId: number, transform: VsdTransform): Uint8Array {
 const drawing = readVsdDrawing(input), shape = drawing.pages.find(p => p.id === pageId)?.shapes.find(s => s.id === shapeId);
 if (!shape?.transformRecord || shape.transformRecord.length !== 65) throw new VsdError('missing-or-formula-transform');
 if(shape.kind!=='shape' || shape.parentId!==0xffffffff || shape.masterPageId!==0xffffffff || shape.masterShapeId!==0xffffffff || shape.unsafeTransform) throw new VsdError('dependent-transform');
 const source=shape.transformRecord.block.bytes;
 for(let i=0;i<7;i++) if(source[shape.transformRecord.offset+i*9] !== (i===6?80:64)) throw new VsdError('unsupported-transform-units');
 const values = [transform.pinX,transform.pinY,transform.width,transform.height,transform.localPinX,transform.localPinY,transform.angle];
 const flipX=transform.flipX,flipY=transform.flipY;
 if (!values.every(Number.isFinite) || values[2]! < 0 || values[3]! < 0 || typeof flipX !== 'boolean' || typeof flipY !== 'boolean') throw new VsdError('invalid-transform');
 const replacement = shape.transformRecord.block.bytes.slice(shape.transformRecord.offset, shape.transformRecord.offset + 65), view = new DataView(replacement.buffer);
 for (let i=0;i<7;i++) view.setFloat64(1+i*9, values[i]!, true);
 replacement[63] = +flipX; replacement[64] = +flipY;
 if (replacement.every((value,i)=>value===shape.transformRecord!.block.bytes[shape.transformRecord!.offset+i])) return input.slice();
 const candidate=replaceRecord(input, drawing, shape.transformRecord, replacement);
 const actual=readVsdDrawing(candidate).pages.find(p=>p.id===pageId)?.shapes.find(s=>s.id===shapeId)?.transform;
 if(!actual || [actual.pinX,actual.pinY,actual.width,actual.height,actual.localPinX,actual.localPinY,actual.angle].some((n,i)=>n!==values[i]) || actual.flipX!==flipX || actual.flipY!==flipY) throw new VsdError('transform-edit-verification');
 return candidate;
}
