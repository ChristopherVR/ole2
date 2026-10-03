import { readFileSync } from 'node:fs';
import { expect, it } from 'vitest';
import { unwrapDocBytes, type DocCfbUnwrap } from '../src/ole-document-doc-cfb.js';
import { readDocFib, readFcLcbAt, patchDocFib, type DocFib } from '../src/ole-document-doc-fib.js';

it('keeps legacy constructed FIB/container interfaces usable by public patch helpers', () => {
	const input = new Uint8Array(readFileSync(new URL('./fixtures/ole-word-97.doc', import.meta.url)));
	const parsed = unwrapDocBytes(input)!;
	const { canRewrite: _capability, ...oldContainer } = parsed;
	const container: DocCfbUnwrap = oldContainer;
	const { nFib: _version, nFibBase: _base, ccpOtherStories: _stories,
		fibRgFcLcbCount: _count, ...oldFib } = readDocFib(container.wordDocBytes);
	const fib: DocFib = oldFib;
	expect(readFcLcbAt(container.wordDocBytes, fib, 33)).toEqual(fib.clx);
	expect(() => readFcLcbAt(container.wordDocBytes, fib, 999)).toThrow();
	const bytes = container.wordDocBytes.slice();
	patchDocFib(bytes, fib, { ccpText: fib.ccpText, cbMac: fib.cbMac, clx: fib.clx,
		plcfbteChpx: fib.plcfbteChpx, plcfbtePapx: fib.plcfbtePapx });
	expect(readDocFib(bytes).clx).toEqual(fib.clx);
});
