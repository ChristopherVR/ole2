import type { WMasterTextStyles } from './write-model.js';

/** Collect master text fonts for the binary FontCollection writer. */
export function masterStyleFonts(styles: WMasterTextStyles | undefined): string[] {
	if (!styles) return [];
	return [styles.title, styles.body, styles.other]
		.flatMap((levels) => levels ?? [])
		.map((level) => level.run.fontName)
		.filter((name): name is string => Boolean(name));
}
