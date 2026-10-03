import {decodeVsdBlock} from './vsd-compression.js';

export const VSD_COMPRESSION_FIT_MAX_BYTES=8*1024*1024;
const MAX_INPUT=VSD_COMPRESSION_FIT_MAX_BYTES,MAX_PADDING=4096,MAX_PROBES=16*1024*1024;
export class VsdCompressionFitError extends Error {
 constructor(readonly reason:'input-budget'|'work-limit'|'verification'){super(`VSD compression ${reason}`);this.name='VsdCompressionFitError';}
}
/** Independently authored bounded LZSS encoding for one existing VSD allocation.
 * The window/address/length representation matches decodeVsdBlock. Refuse any
 * allocation overflow; no page, pointer table or file offset is relocated.
 * Zero suffix padding is reported explicitly for chunk-stream validation. */
export function encodeVsdBlockToSize(input:Uint8Array,capacity:number):{bytes:Uint8Array;paddingBytes:number}|undefined {
 if(!(input instanceof Uint8Array)||input.length>MAX_INPUT||!Number.isSafeInteger(capacity)||capacity<0||capacity>MAX_INPUT)throw new VsdCompressionFitError('input-budget');
 const output=new Uint8Array(capacity),heads=new Int32Array(65536).fill(-1),previous=new Int32Array(input.length).fill(-1);
 let cursor=0,flagOffset=-1,bit=8,tokens=0,probes=0;
 function emit(literal:boolean,a:number,b?:number):boolean{
  const size=(bit===8?1:0)+(literal?1:2);if(size>capacity-cursor)return false;
  if(bit===8){flagOffset=cursor++;bit=0;}
  if(literal)output[flagOffset]!|=1<<bit;
  output[cursor++]=a;if(!literal)output[cursor++]=b!;bit++;tokens++;return true;
 }
 const hash=(p:number)=>((input[p]!*251+input[p+1]!)*251+input[p+2]!)&65535;
 function remember(p:number):void{if(p+2<input.length){const h=hash(p);previous[p]=heads[h]!;heads[h]=p;}}
 function match(position:number):{length:number;source:number}{
  let length=0,source=-1;
  // The not-yet-written final eighteen dictionary cells are initially zero.
  if(position<4060){while(length<18&&position+length<input.length&&input[position+length]===0)length++;if(length>=3)source=-18;else length=0;}
  if(position+2<input.length){
   let candidate=heads[hash(position)]!,attempts=0;
   while(candidate>=Math.max(0,position-4096)&&candidate<position&&attempts++<512){
    let n=0;while(n<18&&position+n<input.length){if(++probes>MAX_PROBES)throw new VsdCompressionFitError('work-limit');if(input[candidate+n]!==input[position+n])break;n++;}
    if(n>=3&&n>length){length=n;source=candidate;if(n===18)break;}candidate=previous[candidate]!;
   }
  }
  return {length,source};
 }
 for(let position=0;position<input.length;){
  let {length,source}=match(position);
  // One literal can unlock a longer next match; this bounded look-ahead avoids
  // short references that make native allocations needlessly overflow.
  if(length>=3&&length<18&&position+1<input.length&&match(position+1).length>length+1)length=0;
  if(length>=3){const address=(source-18)&4095;if(!emit(false,address&255,((address>>4)&240)|(length-3)))return undefined;}
  else{length=1;if(!emit(true,input[position]!))return undefined;}
  for(let i=0;i<length;i++)remember(position+i);position+=length;
 }
 let paddingBytes=0;
 const remaining=capacity-cursor;
 if(remaining){
  // Choose a token count with exact flag/payload storage; the first three
  // zero literals establish a safe window source for any zero references.
  let count=-1,references=0;
  for(let n=1;n<=Math.min(remaining,MAX_PADDING);n++){
   const flags=Math.ceil((tokens+n)/8)-Math.ceil(tokens/8),extra=n+flags,r=remaining-extra;
   if(r>=0&&r<=Math.max(0,n-3)&&n+2*r<=MAX_PADDING){count=n;references=r;break;}
  }
  if(count<0)return undefined;
  for(let i=0;i<count;i++){
   if(i>=3&&references>0){const address=(input.length+paddingBytes-3-18)&4095;if(!emit(false,address&255,((address>>4)&240)))return undefined;paddingBytes+=3;references--;}
   else{if(!emit(true,0))return undefined;paddingBytes++;}
  }
 }
 if(cursor!==capacity)return undefined;
 const decoded=decodeVsdBlock(output,input.length+paddingBytes);
 if(decoded.length!==input.length+paddingBytes||!input.every((v,i)=>decoded[i]===v)||decoded.subarray(input.length).some(v=>v!==0))throw new VsdCompressionFitError('verification');
 return {bytes:output,paddingBytes};
}
