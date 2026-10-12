/** Unified, checked format dispatch. Erased generics never substitute validation. */
import { Ole2DocumentError } from './ole2-document-base.js';
import { readCompoundFileStream } from './ole2-stream-edit.js';
import { CfbDocument } from './cfb-document.js';
import { DocDocument } from './doc-document.js';
import { XlsDocument } from './xls-document.js';
import { PptDocument } from './ppt-document.js';
import { VsdDocument } from './vsd-document.js';

export type Ole2Input = Uint8Array | ArrayBuffer;
export type Ole2Document = CfbDocument | DocDocument | XlsDocument | PptDocument | VsdDocument;
export interface Ole2DocumentsByKind {cfb:CfbDocument;doc:DocDocument;xls:XlsDocument;ppt:PptDocument;vsd:VsdDocument}
export type Ole2Format = keyof Ole2DocumentsByKind;
export interface Ole2ParseOptions<K extends Ole2Format> {readonly expect:K}
export interface Ole2ModelDiagnostic {readonly code:'ambiguous-format'|'unsupported-model';readonly message:string}

export { CfbStorage, CfbStream, type CfbEntry } from './cfb-storage.js';
export { CfbDocument } from './cfb-document.js';

function bytesOf(input:Ole2Input):Uint8Array {
 if(ArrayBuffer.isView(input))return new Uint8Array(new Uint8Array(input.buffer,input.byteOffset,input.byteLength));
 try {return new Uint8Array(ArrayBuffer.prototype.slice.call(input,0));}
 catch {throw new TypeError('OLE2 input must be an ArrayBuffer or Uint8Array');}
}
function detect(bytes:Uint8Array):{format?:'doc'|'xls'|'ppt'|'vsd';ambiguous:boolean} {
 const root=(name:string)=>readCompoundFileStream(bytes,[name])!==undefined;
 const formats:Array<'doc'|'xls'|'ppt'|'vsd'>=[];
 if(root('WordDocument'))formats.push('doc');
 const workbook=root('Workbook'),book=root('Book');
 if(workbook||book)formats.push('xls');
 if(root('PowerPoint Document'))formats.push('ppt');
 if(root('VisioDocument'))formats.push('vsd');
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
   case'vsd':return new VsdDocument(bytes);
  }
 } catch(error) {
  if(options)throw new Ole2DocumentError('unsupported-model',`Cannot decode ${detected.format} model: ${error instanceof Error?error.message:String(error)}`);
  return new CfbDocument(bytes,detected.format,[{code:'unsupported-model',message:error instanceof Error?error.message:String(error)}]);
 }
}
export function parseDoc(input:Ole2Input):DocDocument {return parseOle2(input,{expect:'doc'});}
export function parseXls(input:Ole2Input):XlsDocument {return parseOle2(input,{expect:'xls'});}
export function parsePpt(input:Ole2Input):PptDocument {return parseOle2(input,{expect:'ppt'});}
export function parseVsd(input:Ole2Input):VsdDocument {return parseOle2(input,{expect:'vsd'});}
