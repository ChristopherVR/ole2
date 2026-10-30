/** Shared types and limits for the DOC object readers ([MS-DOC] pictures, shapes,
 * fields, bookmarks and embedded objects). Positions are absolute character
 * positions (CPs) across all stories, as in the piece table. */

export type DocStoryKind = 'main' | 'footnotes' | 'headers' | 'comments' | 'endnotes' | 'textboxes' | 'headerTextboxes';

/** One story's absolute CP range (cpEnd exclusive). */
export interface DocStoryRange {
	readonly story: DocStoryKind;
	readonly cpStart: number;
	readonly cpEnd: number;
}

/** Processing budget; every table is validated against these before allocation. */
export interface DocObjectsLimits {
	/** Characters decoded across all stories. Default 16,777,216. */
	maxCharacters?: number;
	/** Piece table entries. Default 65,536. */
	maxPieces?: number;
	/** Records per table (fields, bookmarks, shapes, BLIPs, pictures, objects). Default 65,536. */
	maxRecords?: number;
	/** Largest single image (compressed or decompressed). Default 67,108,864. */
	maxBlipBytes?: number;
	/** Total image bytes returned. Default 268,435,456. */
	maxTotalBlipBytes?: number;
	/** OfficeArt container nesting depth. Default 16. */
	maxArtDepth?: number;
}

export const DEFAULT_DOC_OBJECTS_LIMITS: Readonly<Required<DocObjectsLimits>> = Object.freeze({
	maxCharacters: 16_777_216,
	maxPieces: 65_536,
	maxRecords: 65_536,
	maxBlipBytes: 67_108_864,
	maxTotalBlipBytes: 268_435_456,
	maxArtDepth: 16,
});

export type DocBlipFormat = 'png' | 'jpeg' | 'bmp' | 'emf' | 'wmf' | 'pict' | 'tiff';

/** A decoded OfficeArt picture. EMF/WMF/PICT bytes are the raw (decompressed) metafile. */
export interface DocBlip {
	/** One-based index in the document BStore, or the one-based position in an inline picture's file-block list. */
	readonly index: number;
	readonly format: DocBlipFormat;
	readonly mime: string;
	/** Usable file bytes: PNG/JPEG/TIFF as stored, DIB wrapped as a BMP file, metafiles decompressed. */
	readonly data: Uint8Array;
	/** Pixel size for bitmaps when the header is readable. */
	readonly widthPx: number | undefined;
	readonly heightPx: number | undefined;
	/** Hex of the 16-byte BLIP unique id. */
	readonly uid: string;
	/** Metafile bounds (EMU) and size (EMU) from the metafile header. */
	readonly metafile?: { readonly boundsEmu: readonly [number, number, number, number]; readonly sizeEmu: readonly [number, number] };
	/** The BLIP was stored out of line (WordDocument "delay" stream). */
	readonly delayed: boolean;
}

/** Crop as fractions (0..1) of the source image. */
export interface DocCrop { readonly top: number; readonly bottom: number; readonly left: number; readonly right: number }

export interface DocInlinePicture {
	/** Absolute CP of the 0x01 picture character. */
	readonly cp: number;
	readonly story: DocStoryKind;
	/** Byte offset of the PICFAndOfficeArtData in the Data stream. */
	readonly dataOffset: number;
	/** Intrinsic size in twips (dxaGoal, dyaGoal) and scale in 1/1000 (mx, my). */
	readonly dxaGoal: number;
	readonly dyaGoal: number;
	readonly mx: number;
	readonly my: number;
	/** Display size in twips: goal * scale / 1000. */
	readonly widthTwips: number;
	readonly heightTwips: number;
	readonly crop: DocCrop | undefined;
	readonly blip: DocBlip | undefined;
	readonly name: string | undefined;
	/** Alternative text from the OfficeArt properties, when present. */
	readonly alt: string | undefined;
}

export type DocShapeWrap = 'topAndBottom' | 'square' | 'none' | 'tight' | 'through' | 'unknown';
export type DocShapeAnchorRelative = 'margin' | 'page' | 'column' | 'paragraph' | 'unknown';

export interface DocFloatingShape {
	readonly spid: number;
	/** Absolute CP of the anchor character. */
	readonly cp: number;
	readonly story: DocStoryKind;
	/** FSPA rectangle in twips, relative to {@link relativeX} and {@link relativeY}. */
	readonly leftTwips: number;
	readonly topTwips: number;
	readonly rightTwips: number;
	readonly bottomTwips: number;
	readonly widthTwips: number;
	readonly heightTwips: number;
	readonly relativeX: DocShapeAnchorRelative;
	readonly relativeY: DocShapeAnchorRelative;
	readonly wrap: DocShapeWrap;
	readonly wrapSide: 'both' | 'left' | 'right' | 'largest' | 'unknown';
	readonly behindText: boolean;
	readonly inHeader: boolean;
	/** Number of text boxes chained from this shape (FSPA cTxbx). */
	readonly textboxCount: number;
	/** OfficeArt shape type (msosptRect = 1, msosptPictureFrame = 75, msosptTextBox = 202, ...). */
	readonly shapeType: number | undefined;
	readonly isPicture: boolean;
	readonly blip: DocBlip | undefined;
	readonly blipIndex: number | undefined;
	readonly crop: DocCrop | undefined;
	/** Text-box id (OfficeArt lTxid): high 16 bits index the text-box PLC (one-based). */
	readonly textboxId: number | undefined;
	readonly fill: { readonly enabled: boolean; readonly colorRgb: number | undefined; readonly opacity: number | undefined } | undefined;
	readonly line: { readonly enabled: boolean; readonly colorRgb: number | undefined; readonly widthEmu: number | undefined } | undefined;
	readonly rotationDegrees: number | undefined;
	readonly name: string | undefined;
	readonly alt: string | undefined;
	/** OfficeArt property table, id (with flag bits cleared) to simple value. */
	readonly properties: ReadonlyMap<number, number>;
}

export type DocFieldKind =
	| 'hyperlink' | 'ref' | 'pageref' | 'toc' | 'page' | 'numpages' | 'date' | 'time' | 'seq'
	| 'mergefield' | 'includepicture' | 'other';

export interface DocFieldBase {
	readonly story: DocStoryKind;
	/** Absolute CPs of the begin (0x13), separator (0x14, if any) and end (0x15) marks. */
	readonly cpBegin: number;
	readonly cpSeparator: number | undefined;
	readonly cpEnd: number;
	/** Field type code ([MS-DOC] flt) stored on the begin mark. */
	readonly flt: number;
	/** Index in the returned list of the enclosing field, or -1. */
	readonly parent: number;
	readonly depth: number;
	/** Instruction text (nested fields replaced by their result), trimmed. */
	readonly instruction: string;
	readonly result: string;
	readonly resultRange: { readonly cpStart: number; readonly cpEnd: number } | undefined;
	readonly kind: DocFieldKind;
	readonly switches: Readonly<Record<string, string | true>>;
	/** Flag byte on the end mark (fDiffer, fZombieEmbed, fResultDirty, fResultEdited, fLocked, fPrivateResult, fNested, fHasSep). */
	readonly endFlags: number;
}

export interface DocHyperlinkField extends DocFieldBase { readonly kind: 'hyperlink'; readonly url: string | undefined; readonly anchor: string | undefined; readonly tooltip: string | undefined; readonly target: string | undefined }
export interface DocRefField extends DocFieldBase { readonly kind: 'ref' | 'pageref'; readonly bookmark: string }
export interface DocTocField extends DocFieldBase { readonly kind: 'toc'; readonly outlineLevels: string | undefined; readonly hyperlinks: boolean; readonly tcEntries: boolean }
export interface DocDateField extends DocFieldBase { readonly kind: 'date' | 'time'; readonly format: string | undefined }
export interface DocSeqField extends DocFieldBase { readonly kind: 'seq'; readonly identifier: string; readonly reset: string | undefined; readonly repeat: boolean }
export interface DocMergeField extends DocFieldBase { readonly kind: 'mergefield'; readonly name: string }
export interface DocIncludePictureField extends DocFieldBase { readonly kind: 'includepicture'; readonly path: string; readonly linked: boolean }
export interface DocPlainField extends DocFieldBase { readonly kind: 'page' | 'numpages' | 'other'; readonly format: string | undefined }
export type DocField = DocHyperlinkField | DocRefField | DocTocField | DocDateField | DocSeqField | DocMergeField | DocIncludePictureField | DocPlainField;

export interface DocBookmark {
	readonly name: string;
	readonly cpStart: number;
	readonly cpEnd: number;
	readonly story: DocStoryKind;
	/** Names starting with an underscore are Word-internal (_Toc, _Ref, _GoBack). */
	readonly hidden: boolean;
	/** Table column range (itcFirst, itcLim exclusive) for bookmarks that cover table columns. */
	readonly columns: { readonly first: number; readonly limit: number } | undefined;
}

export interface DocOleObject {
	/** Absolute CP of the 0x01 object character. */
	readonly cp: number;
	readonly story: DocStoryKind;
	/** ObjectPool storage name, such as `_1234567890`. */
	readonly storageName: string;
	readonly progId: string | undefined;
	/** Class id of the storage, as a lowercase GUID. */
	readonly clsid: string | undefined;
	/** Names of the streams directly inside the storage. */
	readonly streams: readonly string[];
	/** Preview picture, when the run carries one. */
	readonly preview: DocInlinePicture | undefined;
	readonly linked: boolean;
}

export interface DocObjects {
	readonly stories: readonly DocStoryRange[];
	readonly inlinePictures: readonly DocInlinePicture[];
	readonly floatingShapes: readonly DocFloatingShape[];
	readonly fields: readonly DocField[];
	readonly bookmarks: readonly DocBookmark[];
	readonly oleObjects: readonly DocOleObject[];
	/** The document BStore, in order (index = position + 1). */
	readonly blips: readonly DocBlip[];
	/** Non-fatal problems: damaged optional structures are skipped and named here. */
	readonly warnings: readonly string[];
}
