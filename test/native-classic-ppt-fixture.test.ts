import { readFileSync } from 'node:fs';
import { expect, it } from 'vitest';
import { readPptSlideTexts } from '../src/legacy-ppt-text.js';

it('retains the owned classic rectangle text for independently validated geometry cases', () => {
  const bytes = new Uint8Array(readFileSync(new URL('./fixtures/ppt/classic-rectangle.ppt', import.meta.url)));
  const document = readPptSlideTexts(bytes);
  expect(document.slides).toHaveLength(1);
  expect(document.slides[0]!.slideId).toBe(256);
  expect(document.slides[0]!.texts.map(text => text.text)).toEqual(['Owned rectangle']);
  // This regression protects the corpus identity; external PowerPoint snapshots
  // independently establish the declared geometry and character-font behavior.
});
