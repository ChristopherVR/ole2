/** Preservation model for SST/CONTINUE, MS-XLS 2.4.265 and 2.5.293.
 * Parsing is strict and bounded by the serialized table size. Existing entries
 * are never reserialized, so their rich runs, phonetic data and segment layout
 * survive edits byte-for-byte.
 */
import { readRecords, type BiffRecord } from './legacy-excel-biff8.js';

export interface PreservedXlsSst {
 start: number;
 end: number;
 total: number;
 entries: ReadonlyArray<{ text: string; richRuns: number; extendedBytes: number; offset: number; recordOffset: number }>;
 records: readonly BiffRecord[];
}

class SstCursor {
 private segment = 0;
 private offset = 0;
 private remaining: number;
 constructor(private readonly segments: readonly Uint8Array[]) {
  this.remaining = segments.reduce((sum,s)=>sum+s.length,0);
 }
 private normalize(): void {
  while (this.segment < this.segments.length && this.offset === this.segments[this.segment]!.length) { this.segment++; this.offset = 0; }
 }
 get bytesLeft(): number { return this.remaining; }
 location(): {segment:number; offset:number} { this.normalize(); return {segment:this.segment,offset:this.offset}; }
 u8(): number {
  this.normalize();
  if (this.remaining < 1) throw new Error('Truncated SST');
  this.remaining--;
  return this.segments[this.segment]![this.offset++]!;
 }
 u16(): number { return this.u8() | (this.u8()<<8); }
 u32(): number { return (this.u16() | (this.u16()<<16))>>>0; }
 skip(size: number): void {
  if (!Number.isSafeInteger(size) || size < 0 || size > this.remaining) throw new Error('Truncated SST metadata');
  for (let i=0;i<size;i++) this.u8();
 }
 chars(count: number, wide: boolean): string {
  let text = '';
  for (let i=0;i<count;i++) {
   const part = this.segments[this.segment];
   if (!part || this.offset === part.length) {
    this.normalize();
    const flag = this.u8();
    if (flag & 0xfe) throw new Error('Unsupported character continuation flag');
    wide = Boolean(flag & 1);
   }
   if (this.segments[this.segment]!.length-this.offset < (wide?2:1)) throw new Error('Split UTF16 code unit');
   text += String.fromCharCode(wide?this.u16():this.u8());
  }
  return text;
 }
}

/** Validate one complete SST chain in the workbook globals. Throws on malformed
 * counts, characters, metadata or trailing data instead of reading into a cell.
 */
export function readPreservedXlsSst(bytes: Uint8Array): PreservedXlsSst {
 const physical = readRecords(bytes,0,bytes.length);
 const eof = physical.find(r=>r.opcode===0x000a)?.headerOffset ?? -1;
 const tables = physical.filter(r=>r.opcode===0x00fc && r.headerOffset<eof);
 if (tables.length!==1) throw new Error('Expected one SST');
 const sst = tables[0]!;
 const index = physical.indexOf(sst);
 const records: BiffRecord[] = [sst];
 for (let i=index+1; physical[i]?.opcode===0x003c; i++) records.push(physical[i]!);
 if (records.some(r=>r.length===0 || r.length>8224)) throw new Error('Invalid SST record length');
 const cursor = new SstCursor(records.map(r=>bytes.subarray(r.dataOffset,r.dataOffset+r.length)));
 const total = cursor.u32();
 const count = cursor.u32();
 if (count>Math.floor(cursor.bytesLeft/3)) throw new Error('SST entry count exceeds payload');
 const entries: Array<{text:string;richRuns:number;extendedBytes:number;offset:number;recordOffset:number}> = [];
 for (let i=0;i<count;i++) {
  const location = cursor.location();
  const source = records[location.segment]!;
  const offset = source.dataOffset+location.offset;
  const chars = cursor.u16();
  const flags = cursor.u8();
  if (chars>32767 || (flags & 0xf2)) throw new Error('Unsupported SST string');
  const richRuns = flags & 8 ? cursor.u16() : 0;
  const extendedBytes = flags & 4 ? cursor.u32() : 0;
  const text = cursor.chars(chars,Boolean(flags&1));
  cursor.skip(richRuns*4+extendedBytes);
  entries.push({text,richRuns,extendedBytes,offset,recordOffset:source.headerOffset});
 }
 if (cursor.bytesLeft!==0) throw new Error('Unaccounted SST bytes');
 const last = records[records.length-1]!;
 return {start:sst.headerOffset,end:last.dataOffset+last.length,total,entries,records};
}

/** Serialize only an appended plain string as one or more CONTINUE records.
 * A continuation of character data starts with fHighByte; a new string starts
 * with its normal cch/flags header. Every record payload is <=8224 bytes.
 */
export function appendXlsSstString(value: string): Uint8Array {
 if (value.length>32767) throw new Error('XLS strings are limited to 32767 UTF16 units');
 const wide = Array.from({length:value.length},(_,i)=>value.charCodeAt(i)).some(c=>c>255);
 const records: number[] = [];
 let index = 0;
 let first = true;
 do {
  const data = first ? [value.length&255,value.length>>>8,wide?1:0] : [wide?1:0];
  const capacity = Math.floor((8224-data.length)/(wide?2:1));
  let end = Math.min(value.length,index+capacity);
  if (wide && end<value.length && end>index && value.charCodeAt(end-1)>=0xd800 && value.charCodeAt(end-1)<=0xdbff && value.charCodeAt(end)>=0xdc00 && value.charCodeAt(end)<=0xdfff) end--;
  for (;index<end;index++) { const c=value.charCodeAt(index); data.push(c&255); if(wide) data.push(c>>>8); }
  records.push(0x3c,0,data.length&255,data.length>>>8,...data);
  first=false;
 } while(index<value.length);
 return Uint8Array.from(records);
}
