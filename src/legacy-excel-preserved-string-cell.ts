/** Public preservation-safe BIFF8 string edit for bare streams and CFB files. */
import { OLE_MAGIC } from './ole2-parser-types.js';
import { readCompoundFileStream, replaceCompoundFileStream } from './ole2-stream-edit.js';
import { resizeCompoundFileStream } from './ole2-stream-resize.js';
import { editXlsStringWorkbookStream, type XlsPreservedStringResult } from './legacy-excel-preserved-string-edit.js';
export type { XlsPreservedStringResult, XlsPreservedStringFailure } from './legacy-excel-preserved-string-edit.js';

/** Replace an existing LABELSST/RK/NUMBER/MULRK cell with plain Unicode text while retaining
 * the complete serialized workbook and every unrelated CFB stream/allocation.
 * New SST entries support CONTINUE; worksheetIndex follows worksheet tab order.
 * Formula caches are preserved and callers must arrange recalculation.
 * CFB resizing is limited to the layouts supported by resizeCompoundFileStream.
 */
export function editXlsPreservedStringCell(
 input: Uint8Array,
 edit: {row:number;col:number;value:string;worksheetIndex?:number},
): XlsPreservedStringResult {
 if(!OLE_MAGIC.every((value,index)=>input[index]===value)) return editXlsStringWorkbookStream(input,edit);
 const streams=['Workbook','Book'].map(name=>({name,bytes:readCompoundFileStream(input,[name])})).filter(s=>s.bytes!==undefined);
 if(streams.length!==1) return {status:'unchanged',bytes:input,reason:'unsupported-workbook'};
 const stream=streams[0]!;
 const result=editXlsStringWorkbookStream(stream.bytes!,edit);
 if(result.status==='unchanged') return {...result,bytes:input};
 if(result.bytes.length===stream.bytes!.length) {
  const bytes=replaceCompoundFileStream(input,[stream.name],result.bytes);
  return bytes===input ? {status:'unchanged',bytes:input,reason:'container-not-writable'} : {...result,bytes};
 }
 const resized=resizeCompoundFileStream(input,[stream.name],result.bytes);
 return resized.ok ? {...result,bytes:resized.bytes} : {status:'unchanged',bytes:input,reason:'container-not-writable'};
}
