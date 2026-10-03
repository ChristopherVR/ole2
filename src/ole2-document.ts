/** Unified, checked format dispatch. Erased generics never substitute validation. */
import { Ole2DocumentBase, Ole2DocumentError, UnsupportedOle2EditError, type Ole2DocumentCapabilities } from './ole2-document-base.js';
import { readCompoundFileStream } from './ole2-stream-edit.js';
import { resizeCompoundFileStream } from './ole2-stream-resize.js';
import { DocDocument } from './doc-document.js';
import { XlsDocument } from './xls-document.js';
import { PptDocument } from './ppt-document.js';

export type Ole2Input = Uint8Array | ArrayBuffer;
export type Ole2Document = CfbDocument | DocDocument | XlsDocument | PptDocument;
export interface Ole2DocumentsByKind {cfb:CfbDocument;doc:DocDocument;xls:XlsDocument;ppt:PptDocument}
export type Ole2Format = keyof Ole2DocumentsByKind;
export interface Ole2ParseOptions<K extends Ole2Format> {readonly expect:K}
export interface Ole2ModelDiagnostic {readonly code:'ambiguous-format'|'unsupported-model';readonly message:string}

/** Container inspection remains available even when an Office model is unsafe
 * or unsupported. Stream edits repair only CFB allocation, not Office records. */
export class CfbDocument extends Ole2DocumentBase {
 get kind(): 'cfb' {return 'cfb';}
 get capabilities(): Ole2DocumentCapabilities {return CFB_CAPABILITIES;}
 declare readonly detectedFormat: 'doc'|'xls'|'ppt'|undefined;
 declare readonly diagnostics: readonly Ole2ModelDiagnostic[];
 constructor(input:Uint8Array,detectedFormat?:'doc'|'xls'|'ppt',diagnostics:readonly Ole2ModelDiagnostic[]=[]) {
  super(input);
  Object.defineProperty(this,'detectedFormat',{value:detectedFormat,enumerable:true,writable:false});
  Object.defineProperty(this,'diagnostics',{value:Object.freeze(diagnostics.map(item=>Object.freeze({...item}))),enumerable:true,writable:false});
 }
 /** Existing stream by full storage path; no new directory entries are created. */
 stream(path:readonly string[]): CfbStream {
  const stable=Object.freeze([...path]);
  if(!readCompoundFileStream(this.getBytes(),stable))throw new UnsupportedOle2EditError('stream-not-found');
  const stream=new CfbStream(stable,()=>readCompoundFileStream(this.getBytes(),stable),bytes=>{
   const result=resizeCompoundFileStream(this.getBytes(),stable,bytes);
   if(!result.ok)throw new UnsupportedOle2EditError(result.reason);
   this.commitBytes(result.bytes);
  });
  Object.freeze(stream);
  return stream;
 }
}
const CFB_CAPABILITIES=Object.freeze({
 read:Object.freeze(['directory','stream-by-path']),
 write:Object.freeze(['existing-stream-bytes']),
 limitations:Object.freeze(['Container editing does not update Office format records.','New directory entries and external DIFAT expansion are unsupported.','Variable-length v4 mini-stream edits are unsupported.']),
});
export class CfbStream {
 constructor(readonly path:readonly string[],private readonly read:()=>Uint8Array|undefined,private readonly write:(bytes:Uint8Array)=>void) {}
 get bytes():Uint8Array {const value=this.read();if(!value)throw new UnsupportedOle2EditError('stale-stream');return new Uint8Array(value);}
 set bytes(value:Uint8Array) {if(!(value instanceof Uint8Array))throw new UnsupportedOle2EditError('invalid-stream-bytes');this.write(new Uint8Array(value));}
}

function bytesOf(input:Ole2Input):Uint8Array {
 if(ArrayBuffer.isView(input))return new Uint8Array(new Uint8Array(input.buffer,input.byteOffset,input.byteLength));
 try {return new Uint8Array(ArrayBuffer.prototype.slice.call(input,0));}
 catch {throw new TypeError('OLE2 input must be an ArrayBuffer or Uint8Array');}
}
function detect(bytes:Uint8Array):{format?:'doc'|'xls'|'ppt';ambiguous:boolean} {
 const root=(name:string)=>readCompoundFileStream(bytes,[name])!==undefined;
 const formats:Array<'doc'|'xls'|'ppt'>=[];
 if(root('WordDocument'))formats.push('doc');
 const workbook=root('Workbook'),book=root('Book');
 if(workbook||book)formats.push('xls');
 if(root('PowerPoint Document'))formats.push('ppt');
 return {format:formats.length===1?formats[0]:undefined,ambiguous:formats.length>1||(workbook&&book)};
}
/** Default parsing preserves the original CFB inspection surface. Unsupported
 * Office decoding falls back to CfbDocument with an explicit diagnostic.
 * Checked parsing requires a runtime expectation. `expect:'cfb'` explicitly
 * selects container inspection for any valid compound file. */
export function parseOle2(input:Ole2Input):Ole2Document;
export function parseOle2<K extends Ole2Format>(input:Ole2Input,options:Ole2ParseOptions<K>):Ole2DocumentsByKind[K];
export function parseOle2(input:Ole2Input,options?:Ole2ParseOptions<Ole2Format>):Ole2Document {
 const bytes=bytesOf(input);
 const container=new CfbDocument(bytes);
 const detected=detect(bytes);
 if(options?.expect==='cfb')return new CfbDocument(bytes,detected.format,detected.ambiguous?[{code:'ambiguous-format',message:'Multiple root format identities; container inspection only.'}]:[]);
 if(detected.ambiguous) {
  if(options)throw new Ole2DocumentError('ambiguous-format','Multiple root format identities; no format model can be selected.');
  return new CfbDocument(bytes,undefined,[{code:'ambiguous-format',message:'Multiple root format identities; container inspection only.'}]);
 }
 if(options&&options.expect!==detected.format)throw new Ole2DocumentError('format-mismatch',`Expected ${options.expect}; detected ${detected.format??'cfb'}.`);
 if(!detected.format)return container;
 try {
  switch(detected.format) {
   case'doc':return new DocDocument(bytes);
   case'xls':return new XlsDocument(bytes);
   case'ppt':return new PptDocument(bytes);
  }
 } catch(error) {
  if(options)throw new Ole2DocumentError('unsupported-model',`Cannot decode ${detected.format} model: ${error instanceof Error?error.message:String(error)}`);
  return new CfbDocument(bytes,detected.format,[{code:'unsupported-model',message:error instanceof Error?error.message:String(error)}]);
 }
}
export function parseDoc(input:Ole2Input):DocDocument {return parseOle2(input,{expect:'doc'});}
export function parseXls(input:Ole2Input):XlsDocument {return parseOle2(input,{expect:'xls'});}
export function parsePpt(input:Ole2Input):PptDocument {return parseOle2(input,{expect:'ppt'});}
