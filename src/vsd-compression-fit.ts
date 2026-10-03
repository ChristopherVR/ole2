import {decodeVsdBlock} from './vsd-compression.js';

export const VSD_COMPRESSION_FIT_MAX_BYTES=8*1024*1024;
const MAX_INPUT=VSD_COMPRESSION_FIT_MAX_BYTES,MAX_GAP=4096,MAX_PROBES=16*1024*1024;
export class VsdCompressionFitError extends Error {
 constructor(readonly reason:'input-budget'|'work-limit'|'verification'){super(`VSD compression ${reason}`);this.name='VsdCompressionFitError';}
}
/** Independently authored bounded LZSS encoding for one existing VSD allocation.
 * The window/address/length representation matches decodeVsdBlock. Refuse any
 * allocation overflow; no page, pointer table or file offset is relocated.
 * Reference tokens may be split into literals to fit exactly, without adding decoded bytes. */
export function encodeVsdBlockToSize(input:Uint8Array,capacity:number):{bytes:Uint8Array;paddingBytes:number}|undefined {
 if(!(input instanceof Uint8Array)||input.length>MAX_INPUT||!Number.isSafeInteger(capacity)||capacity<0||capacity>MAX_INPUT)throw new VsdCompressionFitError('input-budget');
 const output=new Uint8Array(capacity),heads=new Int32Array(65536).fill(-1),previous=new Int32Array(input.length).fill(-1);
 let cursor=0,flagOffset=-1,bit=8,tokens=0,probes=0;
 const tailSize=4096,positions=new Int32Array(tailSize),lengths=new Uint8Array(tailSize),sources=new Int32Array(tailSize),starts=new Int32Array(tailSize);
 function emit(literal:boolean,a:number,b?:number):boolean{
  const size=(bit===8?1:0)+(literal?1:2);if(size>capacity-cursor)return false;
  if(bit===8){flagOffset=cursor++;output[flagOffset]=0;bit=0;}
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
  const slot=tokens%tailSize;positions[slot]=position;lengths[slot]=length>=3?length:1;sources[slot]=source;starts[slot]=cursor;
  if(length>=3){const address=(source-18)&4095;if(!emit(false,address&255,((address>>4)&240)|(length-3)))return undefined;}
  else{length=1;if(!emit(true,input[position]!))return undefined;}
  for(let i=0;i<length;i++)remember(position+i);position+=length;
 }
 const remaining=capacity-cursor;
 if(remaining){
  if(remaining>MAX_GAP||!tokens)return undefined;
  type Plan={extra:number;count:number;cost:number;parent?:Plan;index:number;k:number};
  const baseTokens=tokens,tailStart=Math.ceil(Math.max(0,tokens-tailSize)/8)*8;
  let states:Plan[]=[{extra:0,count:0,cost:0,index:-1,k:0}],answer:Plan|undefined,planningWork=0,nodes=0,references=0;
  // At most 512 tail references and 64 retained states. A missing exact fit
  // is refused; never compensate by emitting extra decoded zero bytes.
  for(let index=tokens-1;index>=tailStart&&!answer;index--){
   const length=lengths[index%tailSize]!;if(length<3)continue;if(++references>512)break;
   const next=new Map<number,Plan>();for(const state of states)next.set(state.cost*8+(baseTokens+state.count)%8,state);
   for(const state of states){for(let k=1;k<=length;k++){
    if(++planningWork>1024*1024||nodes>=65536)throw new VsdCompressionFitError('work-limit');
    if(k>length-3&&k!==length)continue;
    const count=state.count+(k===length?k-1:k),extra=state.extra+(k===length?k-2:k),cost=extra+Math.ceil((baseTokens+count)/8)-Math.ceil(baseTokens/8);
    if(cost>remaining)continue;const key=cost*8+(baseTokens+count)%8;if(next.has(key))continue;
    const node:Plan={extra,count,cost,parent:state,index,k};nodes++;
    if(cost===remaining){answer=node;break;}next.set(key,node);
   }if(answer)break;}
   states=[...next.values()].sort((a,b)=>b.cost-a.cost).slice(0,63);
   if(!states.some(p=>p.cost===0))states.push({extra:0,count:0,cost:0,index:-1,k:0});
  }
  if(!answer)return undefined;
  const changes=new Map<number,number>();for(let node:Plan|undefined=answer;node?.parent;node=node.parent)changes.set(node.index,node.k);
  cursor=starts[tailStart%tailSize]!;bit=8;flagOffset=-1;tokens=tailStart;
  for(let index=tailStart;index<baseTokens;index++){
   const slot=index%tailSize,position=positions[slot]!,length=lengths[slot]!,k=changes.get(index)??0;
   for(let j=0;j<k;j++)if(!emit(true,input[position+j]!))return undefined;
   if(k<length){if(length===1){if(!emit(true,input[position]!))return undefined;}
    else{const address=(sources[slot]!+k-18)&4095;if(!emit(false,address&255,((address>>4)&240)|(length-k-3)))return undefined;}}
  }
 }
 if(cursor!==capacity)return undefined;
 const decoded=decodeVsdBlock(output,input.length);
 if(decoded.length!==input.length||!input.every((v,i)=>decoded[i]===v))throw new VsdCompressionFitError('verification');
 return {bytes:output,paddingBytes:0};
}
