import { readFileSync } from 'node:fs';
import { describe, expect, it } from 'vitest';
import {
	buildOle2,
	Ole2ParseError,
	parseOle2,
	readOleDocParagraphs,
	writeOleDocParagraphEdit,
} from '../src/index.js';

describe('@office-viewers/ole2 public barrel', () => {
	it('exports CFB APIs that read real PPT and build a readable container', () => {
		const pptBytes = readFileSync(new URL('./fixtures/sample-deck.ppt', import.meta.url));
		const ppt = parseOle2(
			pptBytes.buffer.slice(
				pptBytes.byteOffset,
				pptBytes.byteOffset + pptBytes.byteLength,
			) as ArrayBuffer,
		);
		const stream = ppt.getStream('PowerPoint Document');
		expect(stream).toBeDefined();
		const rebuilt = buildOle2(new Map([['PowerPoint Document', stream!]]));
		expect(parseOle2(rebuilt).getStream('PowerPoint Document')).toEqual(stream);
	});

	it('exports Word paragraph APIs and surfaces malformed OLE input errors', () => {
		const docBytes = new Uint8Array(
			readFileSync(new URL('./fixtures/ole-word-97.doc', import.meta.url)),
		);
		const paragraphs = readOleDocParagraphs(docBytes);
		expect(paragraphs?.length).toBe(4);
		const edited = writeOleDocParagraphEdit(docBytes, 2, 'Public API edit.');
		expect(readOleDocParagraphs(edited)?.[2]).toBe('Public API edit.');
		expect(() => parseOle2(new Uint8Array([1, 2, 3]).buffer)).toThrow(Ole2ParseError);
	});
});
