import { readFileSync } from 'node:fs';
import { describe, expect, it } from 'vitest';
import { readPptSlideTexts, editPptSlideText, PptTextError } from '../src/legacy-ppt-text.js';
import { PptDocument } from '../src/ppt-document.js';
import { readCompoundFileStream, replaceCompoundFileStream } from '../src/ole2-stream-edit.js';
import { resizeCompoundFileStream } from '../src/ole2-stream-resize.js';
import { readRecordOrThrow } from '../src/legacy-ppt-record-stream.js';
import { pptShapeChildren } from '../src/legacy-ppt-shape-reader.js';
import { buildPersistDirectory } from '../src/ppt/persist-directory.js';
import { OA, RT } from '../src/legacy-ppt-record-types.js';
const load = (name = 'ppt/native-text.ppt') => new Uint8Array(readFileSync(new URL(`./fixtures/${name}`, import.meta.url)));
const viewOf = (bytes: Uint8Array) => new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);

/** Duplicate an owned native title atom in one TextHeader group, maintaining framing/save offsets. */
function duplicateTitle(malformedSecondHeader = false): Uint8Array {
  const input = load(), stream = readCompoundFileStream(input, ['PowerPoint Document'])!, user = readCompoundFileStream(input, ['Current User'])!;
  const view = viewOf(stream), oldEditOffset = viewOf(user).getUint32(16, true), chain = buildPersistDirectory(view, oldEditOffset);
  const slideModel = readPptSlideTexts(input).slides[0]!, slide = readRecordOrThrow(view, chain.directory.get(slideModel.persistId)!);
  const text = readRecordOrThrow(view, slideModel.texts[0]!.headerOffset), edit = readRecordOrThrow(view, oldEditOffset);
  const atom = stream.subarray(text.headerOffset, text.dataOffset + text.recLen), tb = new Uint8Array(atom.length + (malformedSecondHeader ? 12 : 0));
  if (malformedSecondHeader) { tb.set(stream.subarray(text.headerOffset - 12, text.headerOffset)); viewOf(tb).setUint16(0, 1 << 4, true); }
  tb.set(atom, malformedSecondHeader ? 12 : 0);
  const sb = stream.subarray(slide.headerOffset, slide.dataOffset + slide.recLen);
  const insertion = text.dataOffset + text.recLen - slide.headerOffset, nextSlide = new Uint8Array(sb.length + tb.length);
  nextSlide.set(sb.subarray(0, insertion)); nextSlide.set(tb, insertion); nextSlide.set(sb.subarray(insertion), insertion + tb.length);
  const ancestors = [slide];
  for (let i = 0; i < ancestors.length; i++) for (const child of pptShapeChildren(view, ancestors[i]!)) {
    if ((child.recVer === 15 || child.recType === OA.ClientTextbox) && child.dataOffset <= text.headerOffset && child.dataOffset + child.recLen >= text.dataOffset + text.recLen) ancestors.push(child);
  }
  for (const parent of ancestors) viewOf(nextSlide).setUint32(parent.headerOffset - slide.headerOffset + 4, parent.recLen + tb.length, true);
  const dirOffset = stream.length + nextSlide.length, editOffset = dirOffset + 16, eb = stream.subarray(edit.headerOffset, edit.dataOffset + edit.recLen);
  const output = new Uint8Array(editOffset + eb.length), ov = viewOf(output);
  output.set(stream); output.set(nextSlide, stream.length); output.set(eb, editOffset);
  ov.setUint16(dirOffset + 2, RT.PersistDirectoryAtom, true); ov.setUint32(dirOffset + 4, 8, true);
  ov.setUint32(dirOffset + 8, (1 << 20) | slideModel.persistId, true); ov.setUint32(dirOffset + 12, stream.length, true);
  ov.setUint32(editOffset + 16, oldEditOffset, true); ov.setUint32(editOffset + 20, dirOffset, true);
  const nextUser = user.slice(); viewOf(nextUser).setUint32(16, editOffset, true);
  const resized = resizeCompoundFileStream(replaceCompoundFileStream(input, ['Current User'], nextUser), ['PowerPoint Document'], output);
  if (!resized.ok) throw new Error(resized.reason); return resized.bytes;
}

describe('bounded PPT text decoding and TextHeader ownership', () => {
  it('counts UTF-16/compressed payload bytes cumulatively and accepts the exact boundary', () => {
    const input = load(), model = readPptSlideTexts(input), stream = readCompoundFileStream(input, ['PowerPoint Document'])!, view = viewOf(stream);
    const total = model.slides.flatMap(slide => slide.texts).reduce((sum, atom) => sum + readRecordOrThrow(view, atom.headerOffset).recLen, 0);
    expect(readPptSlideTexts(input, { maxTextBytes: total })).toEqual(model);
    expect(() => readPptSlideTexts(input, { maxTextBytes: total - 1 })).toThrow(/text bytes/);
    expect(() => readPptSlideTexts(input, { maxRecords: 1 })).toThrow(PptTextError);
  });
  it('bounds slide count before materializing the full model', () => {
    const input = load('sample-deck.ppt');
    expect(readPptSlideTexts(input, { maxSlides: 7 }).slides).toHaveLength(7);
    expect(() => readPptSlideTexts(input, { maxSlides: 6 })).toThrow(/slides/);
  });
  it.each([0, -1, NaN, Infinity, 1.5])('rejects invalid resource limits %s', value => {
    for (const key of ['maxRecords', 'maxSlides', 'maxTextBytes']) expect(() => readPptSlideTexts(load(), { [key]: value })).toThrow(/Invalid text resource limit/);
  });
  it('preserves readable duplicate atoms but refuses edits in their shared TextHeader group atomically', () => {
    const input = duplicateTitle(), model = readPptSlideTexts(input), original = model.slides[0]!.texts[0]!.text;
    expect(model.slides[0]!.texts.slice(0, 2).map(atom => atom.text)).toEqual([original, original]);
    const document = new PptDocument(input), title = document.slides[0]!.texts[0]!;
    title.text = original; expect(document.dirty).toBe(false);
    expect(() => { title.text = original.replace('fixture', 'updated'); }).toThrow(/TextHeader owner/);
    expect(document.revision).toBe(0); expect(title.text).toBe(original); expect(document.serialize()).toEqual(input);
  });
  it('resets valid ownership when a later malformed TextHeader starts another group', () => {
    const input = duplicateTitle(true), model = readPptSlideTexts(input), text = model.slides[0]!.texts[1]!;
    expect(model.slides[0]!.texts[0]!.text).toBe(text.text);
    const result = editPptSlideText(input, { slideIndex: 0, textIndex: 1, expectedText: text.text, text: text.text.replace('fixture', 'updated') });
    expect(result.status).toBe('unsupported'); expect(result.bytes).toBe(input);
    if (result.status === 'unsupported') expect(result.reason).toMatch(/TextHeader owner/);
  });
  it('refuses absent or invalid TextHeader ownership while preserving readable text', () => {
    for (const type of [RT.TextHeaderAtom, 0x7abc]) {
      const input = load(), stream = readCompoundFileStream(input, ['PowerPoint Document'])!.slice(), view = viewOf(stream);
      const model = readPptSlideTexts(input), title = model.slides[0]!.texts[0]!;
      // This native fixture places the four-byte TextHeader immediately before its text atom.
      const header = readRecordOrThrow(view, title.headerOffset - 12); expect(header.recType).toBe(RT.TextHeaderAtom);
      view.setUint16(header.headerOffset + 2, type, true);
      if (type === RT.TextHeaderAtom) view.setUint16(header.headerOffset, 1 << 4, true);
      const malformed = replaceCompoundFileStream(input, ['PowerPoint Document'], stream);
      expect(readPptSlideTexts(malformed).slides[0]!.texts[0]!.text).toBe(title.text);
      const result = editPptSlideText(malformed, { slideIndex: 0, textIndex: 0, expectedText: title.text, text: title.text.replace('fixture', 'updated') });
      expect(result.status).toBe('unsupported'); expect(result.bytes).toBe(malformed);
      if (result.status === 'unsupported') expect(result.reason).toMatch(/TextHeader owner/);
    }
  });
});
