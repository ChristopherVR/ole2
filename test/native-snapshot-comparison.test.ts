import { describe, expect, it } from 'vitest';
// Developer tooling is deliberately outside the published runtime package.
// @ts-expect-error JS tooling has no public declaration file.
import { compareNativeSnapshots } from '../scripts/compare-native-snapshots.mjs';

describe('native semantic differential comparisons', () => {
  const baseline = { paragraphs: [{ text: 'Before\r', bold: false }], shapes: 2 };
  it('requires the declared edit while preserving every other captured field', () => {
    const after = { paragraphs: [{ text: 'After\r', bold: false }], shapes: 2 };
    expect(compareNativeSnapshots(baseline, after).pass).toBe(false);
    expect(compareNativeSnapshots(baseline, after, [{ path: '/paragraphs/0/text', value: 'After\r' }])).toEqual({ pass: true, evidence: 'native-consumer-semantic-snapshot', fullFidelity: false });
    expect(compareNativeSnapshots(baseline, { ...after, shapes: 1 }, [{ path: '/paragraphs/0/text', value: 'After\r' }]).pass).toBe(false);
    expect(baseline.paragraphs[0].text).toBe('Before\r');
  });
  it('rejects nonexistent and prototype paths rather than hiding lost fields', () => {
    for (const path of ['/missing', '/paragraphs/9/text', '/__proto__/polluted']) {
      expect(() => compareNativeSnapshots(baseline, baseline, [{ path, value: true }])).toThrow();
    }
  });
});
