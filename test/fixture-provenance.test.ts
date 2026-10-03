import { readFileSync } from 'node:fs';
import { createHash } from 'node:crypto';
import { describe, expect, it } from 'vitest';

const manifest = JSON.parse(readFileSync(new URL('./fixtures/provenance.json', import.meta.url), 'utf8'));
describe('tracked binary fixture provenance', () => {
  for (const fixture of manifest.fixtures) {
    it(`pins ${fixture.path} to its recorded origin`, () => {
      const bytes = readFileSync(new URL(`./fixtures/${fixture.path}`, import.meta.url));
      expect(createHash('sha256').update(bytes).digest('hex')).toBe(fixture.sha256);
      expect(['repository-owned-synthetic', 'unverified-existing-repository-fixture']).toContain(fixture.provenance);
      if (fixture.generator) expect(readFileSync(new URL(`./fixtures/${fixture.generator}`, import.meta.url)).length).toBeGreaterThan(0);
    });
  }
});
