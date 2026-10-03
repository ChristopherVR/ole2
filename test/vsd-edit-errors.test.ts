import { expect, it } from 'vitest';
import { VsdDocument } from '../src/vsd-document.js';
import { UnsupportedOle2EditError } from '../src/ole2-document-base.js';
import { makeVsdFixture } from './fixtures/vsd/generated.js';

it('reports invalid runtime setter types through the shared edit error without committing', () => {
 const doc = new VsdDocument(makeVsdFixture());
 const shape = doc.pages[0]!.shapes[0]!;
 const before = doc.serialize();
 const runtime = shape as unknown as { text: unknown; transform: unknown };
 for (const [set, reason] of [
  [() => { runtime.text = 42; }, 'invalid-text'],
  [() => { runtime.text = null; }, 'invalid-text'],
  [() => { runtime.transform = null; }, 'invalid-transform'],
  [() => { runtime.transform = undefined; }, 'invalid-transform'],
 ] as const) {
  expect(set).toThrow(UnsupportedOle2EditError);
  expect(set).toThrow(`Unsupported OLE2 edit: ${reason}`);
 }
 expect(doc.serialize()).toEqual(before);
 expect(doc.revision).toBe(0);
 expect(doc.dirty).toBe(false);
 expect(shape.text).toBe('Hello\n');
});

it('wraps a throwing transform accessor and preserves the document transaction', () => {
 const doc = new VsdDocument(makeVsdFixture());
 const shape = doc.pages[0]!.shapes[0]!;
 const before = doc.serialize(), transform = shape.transform!;
 let reads = 0;
 const incoming = { ...transform, get width(): number { reads++; throw new Error('caller accessor failure'); } };
 let failure: unknown;
 try { shape.transform = incoming; } catch (error) { failure = error; }
 expect(failure).toBeInstanceOf(UnsupportedOle2EditError);
 expect((failure as UnsupportedOle2EditError).reason).toBe('edit-failed');
 expect(reads).toBe(1);
 expect(doc.serialize()).toEqual(before);
 expect(shape.transform).toEqual(transform);
 expect(doc.revision).toBe(0);
 expect(doc.dirty).toBe(false);
});
