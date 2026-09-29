import { describe, expect, it } from 'vitest';
import { parseOle2 } from '../src/ole2-parser-read.js';
import { buildPptFile } from '../src/legacy-ppt-writer.js';
import type { WDeck } from '../src/legacy-ppt-writer.js';
import { findChild, findChildren, iterateRecords } from '../src/legacy-ppt-record-stream.js';
import type { PptRecord } from '../src/legacy-ppt-record-stream.js';
import { RT } from '../src/legacy-ppt-record-types.js';

const RT_SLIDE_LIST_WITH_TEXT = 0x0ff0;
const RT_SLIDE_PERSIST_ATOM = 0x03f3;
const FIRST_MASTER_ID = 0x80000000;

interface WrittenStructure {
	/** `[persistIdRef, slideId]` of each master-list entry, in order. */
	masterList: Array<[number, number]>;
	masterContainerCount: number;
	/** `SlideAtom.masterIdRef` of each slide, in stream order. */
	slideMasterRefs: number[];
}

async function writeAndRead(deck: WDeck): Promise<WrittenStructure> {
	const bytes = await buildPptFile(deck);
	const cfb = parseOle2(
		bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength) as ArrayBuffer,
	);
	const stream = cfb.getStream('PowerPoint Document')!;
	const view = new DataView(stream.buffer, stream.byteOffset, stream.byteLength);
	const top: PptRecord[] = [...iterateRecords(view, 0, stream.byteLength)];

	const doc = top.find((r) => r.recType === RT.Document)!;
	const masterListRecord = findChildren(view, doc, RT_SLIDE_LIST_WITH_TEXT).find(
		(r) => r.recInstance === 1,
	)!;
	const masterList = findChildren(view, masterListRecord, RT_SLIDE_PERSIST_ATOM).map(
		(atom): [number, number] => [
			view.getUint32(atom.dataOffset, true),
			view.getUint32(atom.dataOffset + 12, true),
		],
	);
	const slideMasterRefs = top
		.filter((r) => r.recType === RT.Slide)
		.map((slide) => {
			const atom = findChild(view, slide, RT.SlideAtom)!;
			// SlideAtom: geom (4) + placeholderId[8] then masterIdRef.
			return view.getUint32(atom.dataOffset + 12, true);
		});
	return {
		masterList,
		masterContainerCount: top.filter((r) => r.recType === RT.MainMaster).length,
		slideMasterRefs,
	};
}

describe('legacy PowerPoint writer: several main masters', () => {
	it('keeps a deck without `masters` on one master that every slide follows', async () => {
		const written = await writeAndRead({
			widthEmu: 9144000,
			heightEmu: 5143500,
			slides: [{ shapes: [] }, { shapes: [] }],
			pictures: [],
		});
		expect(written.masterContainerCount).toBe(1);
		expect(written.masterList).toEqual([[2, FIRST_MASTER_ID]]);
		expect(written.slideMasterRefs).toEqual([FIRST_MASTER_ID, FIRST_MASTER_ID]);
	});

	it('writes every master and points each slide at its own', async () => {
		const written = await writeAndRead({
			widthEmu: 9144000,
			heightEmu: 5143500,
			slides: [
				{ shapes: [], masterIndex: 1 },
				{ shapes: [] },
				{ shapes: [], masterIndex: 2 },
				{ shapes: [], masterIndex: 0 },
			],
			pictures: [],
			masters: [{}, {}, {}],
		});
		expect(written.masterContainerCount).toBe(3);
		expect(written.masterList).toEqual([
			[2, FIRST_MASTER_ID],
			[3, FIRST_MASTER_ID + 1],
			[4, FIRST_MASTER_ID + 2],
		]);
		expect(written.slideMasterRefs).toEqual([
			FIRST_MASTER_ID + 1,
			FIRST_MASTER_ID,
			FIRST_MASTER_ID + 2,
			FIRST_MASTER_ID,
		]);
	});

	it('falls back to the first master for an out-of-range index', async () => {
		const written = await writeAndRead({
			widthEmu: 9144000,
			heightEmu: 5143500,
			slides: [{ shapes: [], masterIndex: 5 }, { shapes: [], masterIndex: -1 }],
			pictures: [],
			masters: [{}, {}],
		});
		expect(written.slideMasterRefs).toEqual([FIRST_MASTER_ID, FIRST_MASTER_ID]);
	});

	it('writes each master with its own text styles', async () => {
		const level = (sizePt: number) => ({ paragraph: {}, run: { sizePt } });
		const single = await buildPptFile({
			widthEmu: 9144000,
			heightEmu: 5143500,
			slides: [{ shapes: [] }],
			pictures: [],
			masters: [{ styles: { title: [level(44)] } }],
		});
		const pair = await buildPptFile({
			widthEmu: 9144000,
			heightEmu: 5143500,
			slides: [{ shapes: [] }, { shapes: [], masterIndex: 1 }],
			pictures: [],
			masters: [{ styles: { title: [level(44)] } }, { styles: { title: [level(28)] } }],
		});
		// The second master is a whole extra MainMaster, not a shared reference.
		expect(pair.length).toBeGreaterThan(single.length);
	});
});
