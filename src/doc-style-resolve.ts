/**
 * Style cascade: resolve the formatting a paragraph or run actually has by
 * walking `basedOn` chains (bounded, cycle safe) and layering document
 * defaults, paragraph style, character style and direct exceptions.
 *
 * @module doc-style-resolve
 */
import { applyCharSprms, cloneCharProps } from './doc-char-props.js';
import type { DocCharProps, DocFontNames } from './doc-char-props.js';
import { applyParaSprms, cloneParaProps } from './doc-para-props.js';
import type { DocParaProps } from './doc-para-props.js';
import type { DocFont } from './doc-fonts.js';
import type { DocStyle, DocStylesheet } from './doc-styles.js';
import type { DocRawSprm } from './doc-sprm.js';

const MAX_DEPTH = 64;

/** Resolved formatting of a paragraph style: its own paragraph properties and the run defaults it implies. */
export interface ResolvedStyle {
	paragraph: DocParaProps;
	character: DocCharProps;
}

/** Cached cascade over one stylesheet. */
export class DocStyleResolver {
	private readonly cache = new Map<number, ResolvedStyle>();
	private readonly root: DocCharProps;

	constructor(
		private readonly sheet: DocStylesheet,
		private readonly fonts: readonly DocFont[],
	) {
		this.root = { sizePt: 10, ...(Object.keys(sheet.defaultFonts).length ? { fontIds: { ...sheet.defaultFonts } } : {}) };
	}

	/** Style chain from the root ancestor down to `istd`; stops on cycles and excessive depth. */
	private chain(istd: number): DocStyle[] {
		const out: DocStyle[] = [];
		const seen = new Set<number>();
		let cur: number | undefined = istd;
		while (cur !== undefined && out.length < MAX_DEPTH && !seen.has(cur)) {
			seen.add(cur);
			const style: DocStyle | undefined = this.sheet.styles[cur];
			if (!style) break;
			out.push(style);
			cur = style.basedOn;
		}
		return out.reverse();
	}

	/** Cascade of a style (paragraph or character) from its ancestors; memoised and shared, never mutate. */
	resolve(istd: number): ResolvedStyle {
		const hit = this.cache.get(istd);
		if (hit) return hit;
		const paragraph: DocParaProps = {};
		const character = cloneCharProps(this.root);
		for (const style of this.chain(istd)) {
			applyParaSprms(paragraph, style.paragraphSprms);
			applyCharSprms(character, style.characterSprms);
		}
		const resolved = { paragraph, character };
		this.cache.set(istd, resolved);
		return resolved;
	}

	/** Base run properties of a paragraph in style `paragraphStyle`, with an optional character style layered on top. */
	baseCharProps(paragraphStyle: number | undefined, characterStyle: number | undefined): DocCharProps {
		const out = cloneCharProps(this.resolve(paragraphStyle ?? 0).character);
		if (characterStyle !== undefined) {
			for (const style of this.chain(characterStyle)) applyCharSprms(out, style.characterSprms);
		}
		return out;
	}

	/** Paragraph properties of style `istd` plus direct sprms. */
	paragraphProps(istd: number, direct: readonly DocRawSprm[]): DocParaProps {
		const out = cloneParaProps(this.resolve(istd).paragraph);
		applyParaSprms(out, direct);
		out.styleId = istd;
		return out;
	}

	/** Fill `fonts` names from `fontIds`. */
	withFontNames(props: DocCharProps): DocCharProps {
		const ids = props.fontIds;
		if (!ids || this.fonts.length === 0) return props;
		const names: DocFontNames = {};
		for (const slot of ['ascii', 'eastAsia', 'hAnsi', 'cs'] as const) {
			const id = ids[slot];
			const name = id === undefined ? undefined : this.fonts[id]?.name;
			if (name) names[slot] = name;
		}
		if (Object.keys(names).length) props.fonts = names;
		return props;
	}
}
