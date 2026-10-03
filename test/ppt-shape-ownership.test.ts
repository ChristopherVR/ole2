import { readFileSync } from 'node:fs';
import { describe, expect, it } from 'vitest';
import { PptDocument } from '../src/ppt-document.js';
import { readCompoundFileStream, replaceCompoundFileStream } from '../src/ole2-stream-edit.js';
import { resizeCompoundFileStream } from '../src/ole2-stream-resize.js';
import { readRecordOrThrow } from '../src/legacy-ppt-record-stream.js';
import { buildPersistDirectory } from '../src/ppt/persist-directory.js';
import { RT } from '../src/legacy-ppt-record-types.js';

const load = () => new Uint8Array(readFileSync(new URL('./fixtures/ppt/native-text.ppt', import.meta.url)));
const viewOf = (b: Uint8Array) => new DataView(b.buffer, b.byteOffset, b.byteLength);
/** Real owned records, wrapped as opaque notes data with a valid incremental save. */
export function nestedSlideInNotes(): Uint8Array {
  const input = load(), stream = readCompoundFileStream(input, ['PowerPoint Document'])!, user = readCompoundFileStream(input, ['Current User'])!;
  const view = viewOf(stream), oldEditOffset = viewOf(user).getUint32(16, true), chain = buildPersistDirectory(view, oldEditOffset);
  const slide = readRecordOrThrow(view, chain.directory.get(15)!), note = readRecordOrThrow(view, chain.directory.get(18)!), edit = readRecordOrThrow(view, oldEditOffset);
  const sb = stream.subarray(slide.headerOffset, slide.dataOffset + slide.recLen), nb = stream.subarray(note.headerOffset, note.dataOffset + note.recLen);
  const eb = stream.subarray(edit.headerOffset, edit.dataOffset + edit.recLen);
  const noteOffset = stream.length, wrapper = noteOffset + nb.length, slideOffset = wrapper + 8, dirOffset = slideOffset + sb.length, editOffset = dirOffset + 24;
  const out = new Uint8Array(editOffset + eb.length), ov = viewOf(out);
  out.set(stream); out.set(nb, noteOffset); out.set(sb, slideOffset);
  ov.setUint32(noteOffset + 4, note.recLen + 8 + sb.length, true);
  ov.setUint16(wrapper + 2, 0x7abc, true); ov.setUint32(wrapper + 4, sb.length, true);
  ov.setUint16(dirOffset + 2, RT.PersistDirectoryAtom, true); ov.setUint32(dirOffset + 4, 16, true);
  ov.setUint32(dirOffset + 8, (1 << 20) | 15, true); ov.setUint32(dirOffset + 12, slideOffset, true);
  ov.setUint32(dirOffset + 16, (1 << 20) | 18, true); ov.setUint32(dirOffset + 20, noteOffset, true);
  out.set(eb, editOffset); ov.setUint32(editOffset + 16, oldEditOffset, true); ov.setUint32(editOffset + 20, dirOffset, true);
  const nextUser = user.slice(); viewOf(nextUser).setUint32(16, editOffset, true);
  const resized = resizeCompoundFileStream(replaceCompoundFileStream(input, ['Current User'], nextUser), ['PowerPoint Document'], out);
  if (!resized.ok) throw new Error(`Owned adversarial fixture allocation failed: ${resized.reason}`);
  return resized.bytes;
}
describe('PPT geometry persist ownership', () => {
  it('refuses an anchor inside an active slide also owned by opaque notes bytes', () => {
    const input = nestedSlideInNotes(), document = new PptDocument(input), shape = document.slides[0]!.shapes[0]!;
    expect(document.slides.every(slide => slide.notesStatus === 'present')).toBe(true);
    expect(shape.x).toBe(240); shape.x = shape.x; expect(document.dirty).toBe(false);
    expect(() => { shape.x += 8; }).toThrow(/overlaps another live persist object/);
    expect(shape.x).toBe(240); expect(document.revision).toBe(0); expect(document.serialize()).toEqual(input);
  });
  it('retains the native-owned exclusive small-anchor edit path', () => {
    const document = new PptDocument(load()), shape = document.slides[0]!.shapes[0]!;
    shape.x += 8; expect(shape.x).toBe(248); expect(document.revision).toBe(1);
    expect(new PptDocument(document.serialize()).slides[0]!.shapes[0]!.x).toBe(248);
  });
  it('refuses slide extents that also contain the active save-directory and user-edit records', () => {
    const input = load(), stream = readCompoundFileStream(input, ['PowerPoint Document'])!, user = readCompoundFileStream(input, ['Current User'])!;
    const view = viewOf(stream), oldOffset = viewOf(user).getUint32(16, true), chain = buildPersistDirectory(view, oldOffset);
    const slide = readRecordOrThrow(view, chain.directory.get(15)!), edit = readRecordOrThrow(view, oldOffset);
    const sb = stream.subarray(slide.headerOffset, slide.dataOffset + slide.recLen), eb = stream.subarray(edit.headerOffset, edit.dataOffset + edit.recLen);
    const dirOffset = stream.length + sb.length, editOffset = dirOffset + 16, out = new Uint8Array(editOffset + eb.length), ov = viewOf(out);
    out.set(stream); out.set(sb, stream.length); out.set(eb, editOffset);
    ov.setUint32(stream.length + 4, out.length - stream.length - 8, true);
    ov.setUint16(dirOffset + 2, RT.PersistDirectoryAtom, true); ov.setUint32(dirOffset + 4, 8, true);
    ov.setUint32(dirOffset + 8, (1 << 20) | 15, true); ov.setUint32(dirOffset + 12, stream.length, true);
    ov.setUint32(editOffset + 16, oldOffset, true); ov.setUint32(editOffset + 20, dirOffset, true);
    const nu = user.slice(); viewOf(nu).setUint32(16, editOffset, true);
    const resized = resizeCompoundFileStream(replaceCompoundFileStream(input, ['Current User'], nu), ['PowerPoint Document'], out);
    expect(resized.ok).toBe(true); if (!resized.ok) return;
    const document = new PptDocument(resized.bytes), shape = document.slides[0]!.shapes[0]!;
    shape.x = shape.x; expect(document.dirty).toBe(false);
    expect(() => { shape.x += 8; }).toThrow(/save-history metadata/);
    expect(shape.x).toBe(240); expect(document.revision).toBe(0); expect(document.serialize()).toEqual(resized.bytes);
  });
});
