/**
 * Binary PowerPoint 97–2003 writer API.
 * Extracted from ChristopherVR/pptx-viewer `packages/core/src/core/ppt/writer`
 * under Apache-2.0; consumer-specific PptxElement conversion remains there.
 */
export { buildPptFile } from './ppt/writer/write-ppt.js';
export type { BuildPptOptions } from './ppt/writer/write-ppt.js';
export { masterStyleFonts } from './ppt/writer/master-style-fonts.js';
export type * from './ppt/writer/write-model.js';
