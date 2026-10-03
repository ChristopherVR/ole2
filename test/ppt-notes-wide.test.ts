import { readFileSync } from 'node:fs';
import { createHash } from 'node:crypto';
import { describe, expect, it } from 'vitest';
import { PptDocument } from '../src/ppt-document.js';
import { readCompoundFileStream } from '../src/ole2-stream-edit.js';
const load = () => new Uint8Array(readFileSync(new URL('./fixtures/ppt/wide-notes.ppt', import.meta.url)));
const before = '\u6771\u4eac \u03a9 \ud83d\ude00\rWide notes.';
const after = '\u5927\u962a \u03b2 \ud83d\ude00\rWide model.';
const body = (d: PptDocument) => d.slides[1]!.notes!.texts.find(text => text.role === 'body')!;
describe('owned native UTF16 speaker notes slots', () => {
  it('reads the licensed owned fixture, resolved notes identity and exact UTF16/control offsets', () => {
    const input = load(), d = new PptDocument(input), t = body(d);
    expect(createHash('sha256').update(input).digest('hex')).toBe('86905ce486c3fae0d4d63bc3f8187ca3d2ac87b46a599a7c6963d337b189ad42');
    expect(t.text).toBe(before); expect(t.text.length).toBe(19); expect(t.encoding).toBe('utf16');
    expect(t.placeholderId).toBe(12); expect(t.shapeId).toBe(6147);
    expect([d.slides[1]!.notes!.notesId, d.slides[1]!.notes!.persistId]).toEqual([257, 7]);
    expect(d.serialize()).toEqual(input); expect(d.dirty).toBe(false);
  });
  it('edits Japanese/Greek/plain text while preserving CR, surrogate pair and all bytes outside the atom', () => {
    const input = load(), d = new PptDocument(input), t = body(d); t.text = after;
    const out = d.serialize(), original = readCompoundFileStream(input, ['PowerPoint Document'])!, edited = readCompoundFileStream(out, ['PowerPoint Document'])!;
    const start = t.headerOffset + 8, end = start + before.length * 2;
    expect(edited.subarray(0, start)).toEqual(original.subarray(0, start)); expect(edited.subarray(end)).toEqual(original.subarray(end));
    expect(out.reduce((count, value, index) => count + Number(value !== input[index]), 0)).toBe(8);
    expect(createHash('sha256').update(out).digest('hex')).toBe('01abcbfe1bdf12afa02cbf9cd1e61a661d8d489b56422b2689d7f25e078c8b20');
    expect(body(new PptDocument(out)).text).toBe(after); expect(t.text).toBe(after); expect(d.revision).toBe(1);
  });
  it.each([
    before.replace('\r', '\n'),
    before.replace('\ud83d\ude00', '\ud83dX'),
    before.replace('\ud83d\ude00', 'X\ude00'),
    before.replace('\u03a9', '\0'),
  ])('refuses control changes and malformed UTF16 %j without partial mutation', text => {
    const input = load(), d = new PptDocument(input), t = body(d);
    expect(() => { t.text = text; }).toThrow(/controls|surrogate/);
    expect(t.text).toBe(before); expect(d.revision).toBe(0); expect(d.serialize()).toEqual(input);
  });
  it('rejects object coercion before a reentrant edit can execute', () => {
    const input = load(), d = new PptDocument(input), t = body(d); let calls = 0;
    const value = { toString() { calls++; d.slides[0]!.texts[0]!.text = 'Native title updated'; return after; } };
    expect(() => { t.text = value as unknown as string; }).toThrow(/must be a string/);
    expect(calls).toBe(0); expect(t.text).toBe(before); expect(d.revision).toBe(0); expect(d.serialize()).toEqual(input);
  });
});
