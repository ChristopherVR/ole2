/**
 * Header and footer stories from `PlcfHdd` ([MS-DOC] 2.8.? PlcfHdd, 2.4.1
 * header document). The first six stories hold the footnote and endnote
 * separators; each section then owns six: even header, odd header, even
 * footer, odd footer, first-page header, first-page footer. A zero-length
 * story means the section inherits the previous section's.
 *
 * @module doc-headers-footers
 */
import type { DocSection } from './doc-sections.js';
import type { DocStoryRange, StoryContext } from './doc-story.js';

export type DocHeaderFooterKind = 'header' | 'footer';
export type DocHeaderFooterVariant = 'default' | 'even' | 'first';

export interface DocHeaderFooter {
	readonly section: number;
	readonly kind: DocHeaderFooterKind;
	/** `default` is the odd-page (or only) header; `even` needs facing pages; `first` needs a title page. */
	readonly variant: DocHeaderFooterVariant;
	/** The section's own story, zero-length when it inherits. */
	readonly range: DocStoryRange;
	/** True when the section has no story of its own for this slot. */
	readonly inherited: boolean;
	/** Section whose story applies (this one, or the nearest earlier one with content); undefined when none. */
	readonly sourceSection: number | undefined;
	/** The story that applies after inheritance, or undefined when the slot is empty for this section. */
	readonly effective: DocStoryRange | undefined;
	/** The variant is actually used by the section (first needs titlePage, even needs facing pages). */
	readonly applies: boolean;
}

export interface DocNoteSeparators {
	readonly footnoteSeparator: DocStoryRange;
	readonly footnoteContinuationSeparator: DocStoryRange;
	readonly footnoteContinuationNotice: DocStoryRange;
	readonly endnoteSeparator: DocStoryRange;
	readonly endnoteContinuationSeparator: DocStoryRange;
	readonly endnoteContinuationNotice: DocStoryRange;
}

const SLOTS: ReadonlyArray<readonly [DocHeaderFooterKind, DocHeaderFooterVariant]> = [
	['header', 'even'],
	['header', 'default'],
	['footer', 'even'],
	['footer', 'default'],
	['header', 'first'],
	['footer', 'first'],
];

/** Read header/footer stories for the given sections. */
export function readHeadersFooters(
	ctx: StoryContext,
	sections: readonly DocSection[],
	facingPages: boolean,
): { headersFooters: DocHeaderFooter[]; separators: DocNoteSeparators | undefined } {
	const plc = ctx.plc(11, 0);
	if (!plc) return { headersFooters: [], separators: undefined };
	const story = (i: number): DocStoryRange => ctx.range('headers', plc.cps[i] ?? 0, plc.cps[i + 1] ?? plc.cps[i] ?? 0);
	const available = plc.cps.length - 1;
	const separators =
		available >= 6
			? {
					footnoteSeparator: story(0),
					footnoteContinuationSeparator: story(1),
					footnoteContinuationNotice: story(2),
					endnoteSeparator: story(3),
					endnoteContinuationSeparator: story(4),
					endnoteContinuationNotice: story(5),
				}
			: undefined;
	if (available < 6 + 6 * sections.length) {
		ctx.warnings.push('PlcfHdd has fewer stories than the sections require');
	}
	const out: DocHeaderFooter[] = [];
	const lastWithContent = new Map<number, number>();
	sections.forEach((section, k) => {
		SLOTS.forEach(([kind, variant], j) => {
			const slot = j;
			const idx = 6 + 6 * k + j;
			if (idx >= available) return;
			const range = story(idx);
			const empty = range.end === range.start;
			if (!empty) lastWithContent.set(slot, k);
			const source = empty ? lastWithContent.get(slot) : k;
			const effective = source === undefined ? undefined : out.find((e) => e.section === source && e.kind === kind && e.variant === variant)?.range ?? (source === k ? range : undefined);
			out.push({
				section: k,
				kind,
				variant,
				range,
				inherited: empty && source !== undefined,
				sourceSection: source,
				effective,
				applies: variant === 'first' ? section.titlePage : variant === 'even' ? facingPages : true,
			});
		});
	});
	return { headersFooters: out, separators };
}
