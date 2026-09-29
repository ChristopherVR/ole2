/**
 * Resolves the main masters a deck writes and the ids that tie slides to
 * them.
 *
 * `[MS-PPT]` identifies a main master by the `slideId` of its
 * `SlidePersistAtom` in the master list, and a slide's
 * `SlideAtom.masterIdRef` names its master by that same id (never by the
 * master's persist id). Master slide ids start at `0x80000000`, which is
 * disjoint from the small positive range real slides use. The first
 * master's id, `0x80000000`, is the value PowerPoint writes for a
 * single-master deck; further masters take consecutive ids.
 *
 * @module ppt/writer/master-list
 */

import type { WDeck, WMaster } from './write-model.js';

/** The slide id of the first main master (see the module doc). */
export const FIRST_MASTER_SLIDE_ID = 0x80000000;

/**
 * The deck's main masters in write order. A deck without `masters` writes
 * one master from its legacy `masterStyles` / `master` fields.
 */
export function resolveDeckMasters(deck: WDeck): WMaster[] {
	if (deck.masters && deck.masters.length > 0) {
		return deck.masters;
	}
	return [{ styles: deck.masterStyles, roundTrip: deck.master }];
}

/** The `SlidePersistAtom.slideId` (and `masterIdRef`) of master `index`. */
export function masterSlideId(index: number): number {
	return FIRST_MASTER_SLIDE_ID + index;
}

/** The master index a slide follows, clamped to the masters actually written. */
export function slideMasterIndex(masterIndex: number | undefined, masterCount: number): number {
	if (masterIndex === undefined || !Number.isInteger(masterIndex)) {
		return 0;
	}
	return masterIndex >= 0 && masterIndex < masterCount ? masterIndex : 0;
}
