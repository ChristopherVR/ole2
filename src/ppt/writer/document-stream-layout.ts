/**
 * Lays out every persist object of the "PowerPoint Document" stream
 * (DocumentContainer, MainMaster, one Slide per slide, optional Notes
 * containers) sequentially and records each one's offset, ready for
 * `write-ppt.ts` to append the persist directory / user edit and optionally
 * encrypt.
 *
 * @module ppt/writer/document-stream-layout
 */

import { buildPictureStore } from './bstore-writer.js';
import { ByteWriter } from './byte-writer.js';
import {
	buildDocumentContainer,
	buildSlidePersistAtom,
	MASTER_SLIDE_ID_SENTINEL,
} from './document-writer.js';
import { buildExObjList } from './ex-obj-list-writer.js';
import { HyperlinkCollector } from './hyperlink-writer.js';
import { MASTER_PLACEHOLDER_ORDER } from './master-placeholders-writer.js';
import { masterStyleFonts } from './master-style-fonts.js';
import { buildOtherTextStyle } from './master-text-styles-writer.js';
import { buildSoundCollection, MediaCollector } from './media-writer.js';
import { buildNotesContainer } from './notes-writer.js';
import { buildExOleObjStg, OleCollector } from './ole-writer.js';
import { buildMainMasterContainer, buildSlideContainer } from './slide-writer.js';
import type { WDeck, WMaster } from './write-model.js';

/**
 * `[MS-PPT]` names a main master by the `slideId` of its `SlidePersistAtom`
 * in the master list, and a slide's `SlideAtom.masterIdRef` refers to that
 * id, never to the master's persist id. The first master's id,
 * `0x80000000`, is the value PowerPoint writes for a single-master deck;
 * further masters take consecutive ids.
 */
function masterSlideId(index: number): number {
	return MASTER_SLIDE_ID_SENTINEL + index;
}

/** The deck's main masters in write order; one from the legacy fields when `masters` is absent. */
function resolveDeckMasters(deck: WDeck): WMaster[] {
	return deck.masters && deck.masters.length > 0
		? deck.masters
		: [{ styles: deck.masterStyles, roundTrip: deck.master }];
}

/** A slide's master index, falling back to the first master when absent or out of range. */
function slideMasterIndex(masterIndex: number | undefined, masterCount: number): number {
	return masterIndex !== undefined &&
		Number.isInteger(masterIndex) &&
		masterIndex >= 0 &&
		masterIndex < masterCount
		? masterIndex
		: 0;
}

/** Count every shape (recursively) in one drawing's shape list, +1 for the patriarch. */
function countDrawingShapes(shapes: WDeck['slides'][number]['shapes']): number {
	let count = 1;
	const walk = (list: WDeck['slides'][number]['shapes']): void => {
		for (const shape of list) {
			count++;
			if (shape.kind === 'group') {
				walk(shape.children);
			}
		}
	};
	walk(shapes);
	return count;
}

/** Collect every distinct font name referenced anywhere in the deck. */
function collectFonts(deck: WDeck, masters: WMaster[]): string[] {
	const collect = (shape: WDeck['slides'][number]['shapes'][number]): string[] => {
		if (shape.kind === 'group') {
			return shape.children.flatMap(collect);
		}
		if (shape.kind !== 'shape' || !shape.text) {
			return [];
		}
		return shape.text.paragraphs.flatMap((p) =>
			p.runs.map((r) => r.fontName).filter((n): n is string => Boolean(n)),
		);
	};
	// Font 0 is the default face of every run and master level that names
	// none: the theme's minor (body) face, as PowerPoint's own SaveAs orders it.
	const themeMinor = masters[0]?.roundTrip?.themeFonts?.minor;
	const fonts = Array.from(
		new Set([
			...(themeMinor ? [themeMinor] : []),
			...deck.slides.flatMap((slide) => slide.shapes.flatMap(collect)),
			...masters.flatMap((master) => masterStyleFonts(master.styles)),
		]),
	);
	return fonts.length > 0 ? fonts : ['Calibri'];
}

/** Result of laying out the unencrypted document stream. */
export interface DocumentStreamLayout {
	bytes: ByteWriter;
	offsets: Array<[number, number]>;
	docId: number;
	maxPersistId: number;
	picturesStream: Uint8Array | undefined;
}

/** Persist id of the DocumentContainer. */
export const DOC_ID = 1;
/** Persist id of the first MainMaster; further masters follow it in order. */
export const MASTER_ID = 2;

/** Lay out the full unencrypted "PowerPoint Document" stream content. */
export function layoutDocumentStream(deck: WDeck): DocumentStreamLayout {
	const slideRect = { x: 0, y: 0, w: deck.widthEmu, h: deck.heightEmu };
	const notesRect = { x: 0, y: 0, w: deck.heightEmu, h: deck.widthEmu };

	const masters = resolveDeckMasters(deck);
	const masterIds = masters.map((_, i) => MASTER_ID + i);
	const firstSlideId = MASTER_ID + masters.length;
	const slideIds = deck.slides.map((_, i) => firstSlideId + i);
	let nextId = firstSlideId + deck.slides.length;
	const notesIds = deck.slides.map((slide) => (slide.notesParagraphs?.length ? nextId++ : 0));

	// Drawing ids: 1..M = masters, then the slides in order, then one per
	// slide that has notes. Every drawing in the document needs a distinct id
	// (see `drawing-writer.ts#buildDrawing`'s doc comment).
	const masterDrawingIds = masters.map((_, i) => 1 + i);
	const slideDrawingIds = deck.slides.map((_, i) => 1 + masters.length + i);
	let nextDrawingId = 1 + masters.length + deck.slides.length;
	const notesDrawingIds = deck.slides.map((slide) =>
		slide.notesParagraphs?.length ? nextDrawingId++ : 0,
	);
	const shapesPerDrawing = [
		...masters.map(() => 1 + MASTER_PLACEHOLDER_ORDER.length), // patriarch + placeholders
		...deck.slides.map((slide) => countDrawingShapes(slide.shapes)),
		...deck.slides.filter((s) => s.notesParagraphs?.length).map(() => 2), // patriarch + body placeholder
	];

	const fonts = collectFonts(deck, masters);
	const { dggContainer, picturesStream } = buildPictureStore(deck.pictures, shapesPerDrawing);

	// Document-wide: every hyperlink/click-action target, every OLE embed, and
	// every embedded audio shape anywhere in the deck (shape and text-run
	// level, across every slide, notes page and the master) shares these
	// three collectors, matching real PowerPoint's single document-level
	// ExObjListContainer (see `ex-obj-list-writer.ts`) and (for audio) single
	// document-level SoundCollectionContainer (see `media-writer.ts`).
	const hyperlinks = new HyperlinkCollector();
	const oleEmbeds = new OleCollector(hyperlinks);
	const mediaEmbeds = new MediaCollector(hyperlinks);

	const slideContainers = deck.slides.map((slide, i) =>
		buildSlideContainer(
			slide,
			slideRect,
			masterSlideId(slideMasterIndex(slide.masterIndex, masters.length)),
			notesIds[i]!,
			fonts,
			slideDrawingIds[i]!,
			hyperlinks,
			oleEmbeds,
			mediaEmbeds,
		),
	);
	const notesContainers = deck.slides
		.map((slide, i) =>
			notesIds[i]
				? buildNotesContainer(
						slide.notesParagraphs!,
						notesRect,
						slideIds[i]!,
						fonts,
						notesDrawingIds[i]!,
						hyperlinks,
						oleEmbeds,
						mediaEmbeds,
					)
				: undefined,
		)
		.filter((c): c is Uint8Array => c !== undefined);

	const masterContainers = masters.map((master, i) =>
		buildMainMasterContainer(
			slideRect,
			masterDrawingIds[i]!,
			hyperlinks,
			oleEmbeds,
			mediaEmbeds,
			master.styles,
			fonts,
			master.roundTrip,
		),
	);
	const masterPersistAtoms = masterIds.map((id, i) => buildSlidePersistAtom(id, masterSlideId(i)));
	// flags=4: real (COM-written) files set this bit on a SLIDE's own
	// SlidePersistAtom (never on a master's); see buildSlidePersistAtom's doc.
	const slidePersistAtoms = slideIds.map((id, i) => buildSlidePersistAtom(id, 256 + i, 4));

	// Every OLE embed's ExOleObjStg is its own persist object (referenced by
	// ExOleObjAtom.persistIdRef), allocated AFTER every slide/notes/master
	// container above so every embed anywhere in the deck has already been
	// registered. One entry per embed, appended to the same id sequence notes
	// used.
	const oleStgIds = oleEmbeds.all.map(() => nextId++);
	oleEmbeds.all.forEach((entry, i) => {
		entry.persistIdRef = oleStgIds[i];
	});
	const oleStgRecords = oleEmbeds.all.map((entry) => buildExOleObjStg(entry.storage));

	// Built AFTER every slide/notes/master container above so every
	// hyperlink/OLE/media target referenced anywhere in the deck has already
	// been registered.
	const exObjList = buildExObjList(hyperlinks, oleEmbeds, mediaEmbeds);
	const soundCollection = buildSoundCollection(mediaEmbeds);

	const documentInput = {
		widthEmu: deck.widthEmu,
		heightEmu: deck.heightEmu,
		fonts,
		masterPersistAtoms,
		slidePersistAtoms,
		dggContainer,
		exObjList,
		soundCollection,
		otherTextStyle: buildOtherTextStyle(masters[0]?.styles, (name) => {
			const idx = name ? fonts.indexOf(name) : -1;
			return idx >= 0 ? idx : undefined;
		}),
	};
	const maxPersistId = nextId - 1;
	const contentSizeWithoutPadding =
		buildDocumentContainer(documentInput).length +
		masterContainers.reduce((sum, c) => sum + c.length, 0) +
		slideContainers.reduce((sum, c) => sum + c.length, 0) +
		notesContainers.reduce((sum, c) => sum + c.length, 0) +
		oleStgRecords.reduce((sum, c) => sum + c.length, 0);
	const paddingBytes = ensureMinimumDocumentStreamSize(contentSizeWithoutPadding, maxPersistId);
	const documentContainer = buildDocumentContainer({ ...documentInput, paddingBytes });

	const layout = new ByteWriter();
	const offsets: Array<[number, number]> = [];
	const place = (id: number, bytes: Uint8Array): void => {
		offsets.push([id, layout.size]);
		layout.bytes(bytes);
	};
	place(DOC_ID, documentContainer);
	masterIds.forEach((id, i) => place(id, masterContainers[i]!));
	slideIds.forEach((id, i) => place(id, slideContainers[i]!));
	let notesCursor = 0;
	notesIds.forEach((id) => {
		if (id) {
			place(id, notesContainers[notesCursor++]!);
		}
	});
	oleStgIds.forEach((id, i) => place(id, oleStgRecords[i]!));

	return { bytes: layout, offsets, docId: DOC_ID, maxPersistId, picturesStream };
}

/**
 * How many extra bytes `buildDocumentContainer` needs as a padding `List`
 * record so the finished "PowerPoint Document" stream (this content, plus
 * the `PersistDirectoryAtom` and `UserEditAtom` `write-ppt.ts` appends
 * after it) ends up in a STRICTLY LARGER 512-byte CFB sector count than the
 * 4096-byte (8-sector) "Current User" stream `current-user-writer.ts`
 * always writes.
 *
 * Confirmed required by direct COM testing across a size sweep (both by
 * slide count and by run length, independently): `Presentations.Open`
 * fails whenever "PowerPoint Document" is <= 4096 bytes (tied with or
 * smaller than "Current User"'s sector count) and succeeds once it reaches
 * 4608 bytes (9 sectors) - a boundary unrelated to slide count, persist-id
 * count, or any of this writer's other fixes. Real PowerPoint never needs
 * this: with a full embedded theme, its own "PowerPoint Document" is always
 * orders of magnitude past this threshold already.
 */
function ensureMinimumDocumentStreamSize(
	contentSizeWithoutPadding: number,
	maxPersistId: number,
): number {
	const CFB_SECTOR_SIZE = 512;
	const CURRENT_USER_STREAM_SIZE = 4096;
	const TARGET_MINIMUM = CURRENT_USER_STREAM_SIZE + CFB_SECTOR_SIZE; // 4608: one whole sector past Current User
	// PersistDirectoryAtom (one contiguous run: 8-byte record header + 4-byte
	// group header + 4 bytes per persist id) + UserEditAtom (8-byte record
	// header + 28-byte data = 36 bytes), matching persist-writer.ts exactly.
	const persistTailSize = 8 + 4 + 4 * maxPersistId + 36;
	const projectedTotal = contentSizeWithoutPadding + persistTailSize;
	return Math.max(0, TARGET_MINIMUM - projectedTotal);
}
