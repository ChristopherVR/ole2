import { readFileSync } from 'node:fs';
import { describe, expect, it } from 'vitest';
import { readPptSlideTexts, editPptSlideText } from '../src/legacy-ppt-text.js';

const load = () => new Uint8Array(readFileSync(new URL('./fixtures/ppt/native-text.ppt', import.meta.url)));
describe('owned PowerPoint16 native-save corpus', () => {
  it('reads actual compressed and wide text including a surrogate pair', () => {
    const deck = readPptSlideTexts(load());
    expect(deck.slides.map(slide => slide.texts.map(atom => atom.text))).toEqual([
      ['Native title fixture', 'Synthetic text body.\rLine two.'],
      ['Unicode Ω fixture', '日本語 café 😀'],
    ]);
    expect(deck.slides.map(slide => slide.texts[0].encoding)).toEqual(['compressed-unicode', 'utf16']);
  });
  it('supports exact-length edits in both encodings while changing only target text bytes', () => {
    const original = load();
    const narrow = editPptSlideText(original, { slideIndex: 0, textIndex: 0, expectedText: 'Native title fixture', text: 'Native title updated' });
    expect(narrow.status).toBe('edited');
    const wide = editPptSlideText(narrow.bytes, { slideIndex: 1, textIndex: 0, expectedText: 'Unicode Ω fixture', text: 'Unicode Ω updated' });
    expect(wide.status).toBe('edited');
    expect(wide.bytes.length).toBe(original.length);
    const deck = readPptSlideTexts(wide.bytes);
    expect(deck.slides[0].texts[0].text).toBe('Native title updated');
    expect(deck.slides[1].texts[0].text).toBe('Unicode Ω updated');
    expect(deck.slides[1].texts[1].text).toBe('日本語 café 😀');
    // Independent native semantics are recorded separately; this is a byte-boundary regression.
    expect(original.reduce((n, byte, index) => n + Number(byte !== wide.bytes[index]), 0)).toBe(14);
  });
});
