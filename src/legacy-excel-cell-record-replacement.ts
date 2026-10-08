/** Checked relocation shared by bounded cell-record conversions. */
import { reject, type Model } from './legacy-excel-record-edit-model.js';
import { relocateXlsDbCells, type XlsCellRecordReplacement } from './legacy-excel-cell-record-splices.js';
import { KNOWN_XLS_RECORDS, XLS_FRT_WRAPPER } from './legacy-excel-known-records.js';

/** Private relocation over this call's checked model and generated cell bytes. */
export function applyXlsCellRecordReplacement(input: Uint8Array, m: Model, replacement: XlsCellRecordReplacement): Uint8Array {
 if (!Number.isSafeInteger(replacement.start) || !Number.isSafeInteger(replacement.end) || replacement.start < 0 || replacement.end <= replacement.start || replacement.end > input.length || !(replacement.bytes instanceof Uint8Array)) reject('malformed-records');
 const target = m.records.find(r => r.headerOffset === replacement.start);
 if (!target || target.dataOffset + target.length !== replacement.end || ![0x0201,0x00be,0x0205,0x0203,0x027e,0x00bd].includes(target.opcode)) reject('malformed-records');
 const delta = replacement.bytes.length - (replacement.end - replacement.start);
 if (delta !== 0 && m.records.some(r => !KNOWN_XLS_RECORDS.has(r.opcode) || r.opcode === XLS_FRT_WRAPPER)) reject('unsupported-pointer-record');
 const map = (address: number): number => {
  if (!Number.isSafeInteger(address) || address < 0 || address > input.length || (address > replacement.start && address < replacement.end)) reject('malformed-records');
  return address >= replacement.end ? address + delta : address;
 };
 const out = new Uint8Array(input.length + delta);
 out.set(input.subarray(0, replacement.start)); out.set(replacement.bytes, replacement.start);
 out.set(input.subarray(replacement.end), replacement.start + replacement.bytes.length);
 if (delta !== 0) {
  relocateXlsDbCells(input, out, m.records, replacement, map);
  const source = new DataView(input.buffer, input.byteOffset, input.byteLength), dest = new DataView(out.buffer), byOffset = new Map(m.records.map(r => [r.headerOffset, r]));
  for (const r of m.records) {
   if (r.opcode === 0x0085) {
    if (r.headerOffset >= m.globalsEnd || r.depth !== 1) reject('malformed-records');
    dest.setUint32(map(r.dataOffset), map(source.getUint32(r.dataOffset, true)), true);
   }
   if (r.opcode === 0x020b) {
    if (r.length < 16 || (r.length - 16) % 4 !== 0) reject('malformed-records');
    const owner = m.worksheets.find(sheet => r.depth === 1 && r.headerOffset >= sheet.start && r.headerOffset < sheet.end);
    if (!owner) reject('malformed-records');
    const pointers = new Set<number>();
    for (let p = r.dataOffset + 12; p < r.dataOffset + r.length; p += 4) {
     const address = source.getUint32(p, true);
     const pointed = byOffset.get(address);
     if (pointed?.opcode !== (p === r.dataOffset + 12 ? 0x0055 : 0x00d7) || pointed.depth !== 1 || address < owner.start || address >= owner.end) reject('malformed-records');
     if (pointers.has(address)) reject('malformed-records');
     pointers.add(address);
     dest.setUint32(map(p), map(address), true);
    }
   }
  }
 }
 return out;
}
