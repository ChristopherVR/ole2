/**
 * Public types of the full BIFF8 (`.xls`, Excel 97-2003) workbook reader
 * (`readXlsWorkbook` in `legacy-excel-workbook.ts`).
 *
 * Indices are zero-based throughout (rows, columns, sheets, XFs, fonts).
 * Colours are resolved against the workbook palette where BIFF8 stores a
 * palette index; system colours keep their index with `auto` set.
 *
 * Reference: [MS-XLS] Excel Binary File Format.
 * @module legacy-excel-workbook-types
 */

/** Excel error literals a BIFF8 cell or formula result can hold. */
export type XlsErrorCode =
	| '#NULL!'
	| '#DIV/0!'
	| '#VALUE!'
	| '#REF!'
	| '#NAME?'
	| '#NUM!'
	| '#N/A'
	| '#GETTING_DATA';

/** An error cell value. */
export interface XlsError {
	error: XlsErrorCode;
}

/** A cell value: number (dates are serial numbers), text, boolean, error or empty. */
export type XlsCellValue = number | string | boolean | XlsError | null;

/** An inclusive cell range. */
export interface XlsRange {
	firstRow: number;
	lastRow: number;
	firstCol: number;
	lastCol: number;
}

/** One cell record (value cells, formulas and formatted blanks). */
export interface XlsCell {
	row: number;
	col: number;
	/** Index into {@link XlsWorkbook.xfs}. */
	xf: number;
	/** The value, or the formula's cached result. `null` for a formatted blank. */
	value: XlsCellValue;
	/** Formula text in A1 notation without a leading `=`, when the tokens could be decoded. */
	formula?: string;
	/** True when the cell holds a formula whose tokens could not be decoded to text. */
	formulaUndecoded?: boolean;
	/** The array formula range, on the anchor cell of an array formula only. */
	arrayRange?: XlsRange;
	/** True when the formula came from a shared-formula (`SHRFMLA`) definition. */
	sharedFormula?: boolean;
}

/**
 * A colour reference. `rgb` (`RRGGBB`) is set when the colour resolves to an
 * explicit value (palette entry or an `XFEXT` full colour); `theme`/`tint`
 * when an `XFEXT` names a theme slot; `auto` for system/automatic colours.
 */
export interface XlsColor {
	/** BIFF palette index (0-63 palette, 64 system foreground, 65 system background, 0x7FFF automatic). */
	index?: number;
	rgb?: string;
	theme?: number;
	tint?: number;
	auto?: boolean;
}

export type XlsUnderline = 'none' | 'single' | 'double' | 'singleAccounting' | 'doubleAccounting';

/** A `FONT` record. */
export interface XlsFont {
	name: string;
	/** Size in points. */
	size: number;
	bold: boolean;
	/** Raw font weight (400 normal, 700 bold). */
	weight: number;
	italic: boolean;
	underline: XlsUnderline;
	strike: boolean;
	outline: boolean;
	shadow: boolean;
	script: 'none' | 'superscript' | 'subscript';
	color: XlsColor;
	/** Font family class (0 any, 1 roman, 2 swiss, 3 modern, 4 script, 5 decorative). */
	family: number;
	charset: number;
}

export type XlsBorderStyle =
	| 'thin'
	| 'medium'
	| 'dashed'
	| 'dotted'
	| 'thick'
	| 'double'
	| 'hair'
	| 'mediumDashed'
	| 'dashDot'
	| 'mediumDashDot'
	| 'dashDotDot'
	| 'mediumDashDotDot'
	| 'slantDashDot';

export interface XlsBorderEdge {
	style: XlsBorderStyle;
	color: XlsColor;
}

export interface XlsBorders {
	left?: XlsBorderEdge;
	right?: XlsBorderEdge;
	top?: XlsBorderEdge;
	bottom?: XlsBorderEdge;
	diagonal?: XlsBorderEdge;
	diagonalUp: boolean;
	diagonalDown: boolean;
}

export type XlsFillPattern =
	| 'none'
	| 'solid'
	| 'mediumGray'
	| 'darkGray'
	| 'lightGray'
	| 'darkHorizontal'
	| 'darkVertical'
	| 'darkDown'
	| 'darkUp'
	| 'darkGrid'
	| 'darkTrellis'
	| 'lightHorizontal'
	| 'lightVertical'
	| 'lightDown'
	| 'lightUp'
	| 'lightGrid'
	| 'lightTrellis'
	| 'gray125'
	| 'gray0625';

export interface XlsFill {
	pattern: XlsFillPattern;
	/** Pattern foreground; for `solid` fills this is the cell colour. */
	fg: XlsColor;
	bg: XlsColor;
}

export type XlsHorizontalAlign =
	| 'general'
	| 'left'
	| 'center'
	| 'right'
	| 'fill'
	| 'justify'
	| 'centerContinuous'
	| 'distributed';
export type XlsVerticalAlign = 'top' | 'center' | 'bottom' | 'justify' | 'distributed';

export interface XlsAlignment {
	horizontal: XlsHorizontalAlign;
	vertical: XlsVerticalAlign;
	wrap: boolean;
	shrinkToFit: boolean;
	indent: number;
	/** 0-90 counter-clockwise, 91-180 clockwise, 255 vertical stacked. */
	rotation: number;
	/** 0 context, 1 left-to-right, 2 right-to-left. */
	readingOrder: number;
}

/** A resolved extended format (`XF` record, with `XFEXT` colours applied). */
export interface XlsXf {
	/** Index into {@link XlsWorkbook.fonts} (the BIFF gap at font index 4 is already removed). */
	font: number;
	/** Number format id. */
	numFmtId: number;
	/** The resolved number format code (built-in en-US code or a `FORMAT` record). */
	numFmt: string;
	/** True for a cell style XF, false for a cell XF. */
	isStyle: boolean;
	/** Parent style XF index (cell XFs). */
	parent: number;
	locked: boolean;
	hidden: boolean;
	alignment: XlsAlignment;
	border: XlsBorders;
	fill: XlsFill;
	/** An `XFEXT` text colour that overrides the font colour for this format. */
	fontColor?: XlsColor;
}

/** A `NAME` record (defined name). */
export interface XlsName {
	name: string;
	/** Formula text without `=`, when decodable. */
	formula?: string;
	/** Zero-based sheet index for a sheet-scoped name. */
	localSheet?: number;
	hidden: boolean;
	builtin: boolean;
	/** True for macro/function names. */
	isFunction: boolean;
}

export interface XlsColumnInfo {
	firstCol: number;
	lastCol: number;
	/** Width in characters (the record's 1/256 units divided by 256). */
	width: number;
	xf: number;
	hidden: boolean;
	customWidth: boolean;
	bestFit: boolean;
	outlineLevel: number;
	collapsed: boolean;
}

export interface XlsRowInfo {
	row: number;
	/** Height in points. */
	height: number;
	customHeight: boolean;
	hidden: boolean;
	outlineLevel: number;
	collapsed: boolean;
	/** Row default format, when the row record carries one. */
	xf?: number;
}

export interface XlsSheetView {
	showGridLines: boolean;
	showHeaders: boolean;
	showZeros: boolean;
	showFormulas: boolean;
	rightToLeft: boolean;
	selected: boolean;
	/** Percent (from `SCL`/`WINDOW2`), absent for 100. */
	zoom?: number;
	/** Top-left visible cell. */
	topRow: number;
	leftCol: number;
	/** Frozen panes: rows above / columns left of the split. */
	freeze?: { rows: number; cols: number; topRow: number; leftCol: number };
	/** Unfrozen split position in twips (x) and twips (y), when the window is split. */
	split?: { x: number; y: number };
	activeCell?: { row: number; col: number };
	selection?: XlsRange[];
}

export interface XlsHyperlink {
	range: XlsRange;
	/** External URL or file path. */
	target?: string;
	/** In-workbook location (`Sheet2!A1`) or URL fragment. */
	location?: string;
	display?: string;
	tooltip?: string;
}

export interface XlsComment {
	row: number;
	col: number;
	author: string;
	text: string;
	visible: boolean;
}

export type XlsSheetState = 'visible' | 'hidden' | 'veryHidden';
export type XlsSheetKind = 'worksheet' | 'chart' | 'macro' | 'vba' | 'unknown';

export interface XlsSheet {
	name: string;
	state: XlsSheetState;
	kind: XlsSheetKind;
	/** Cells in record order (row-major within row blocks). */
	cells: XlsCell[];
	merges: XlsRange[];
	columns: XlsColumnInfo[];
	rows: XlsRowInfo[];
	/** `DEFCOLWIDTH`, in characters. */
	defaultColWidth?: number;
	/** `STANDARDWIDTH`, in characters (1/256 precision). */
	standardWidth?: number;
	/** `DEFAULTROWHEIGHT`, in points. */
	defaultRowHeight?: number;
	/** `DIMENSIONS` (used range), when not empty. */
	dimensions?: XlsRange;
	view: XlsSheetView;
	hyperlinks: XlsHyperlink[];
	comments: XlsComment[];
	tabColor?: XlsColor;
	protection?: { passwordHash?: string };
	/** Content this reader saw but does not model (charts, pictures, conditional formats, ...). */
	unsupported: string[];
}

export interface XlsWorkbook {
	biffVersion: 8;
	/** Dates count from 1904-01-01 (`DATEMODE`). */
	date1904: boolean;
	sheets: XlsSheet[];
	fonts: XlsFont[];
	xfs: XlsXf[];
	/** `FORMAT` records by number format id (custom and locale built-ins). */
	numberFormats: Map<number, string>;
	/** The effective 64-entry palette (`RRGGBB`), `PALETTE` overrides applied. */
	palette: string[];
	names: XlsName[];
	/** Zero-based index of the active sheet (`WINDOW1`). */
	activeSheet: number;
	/** `CODEPAGE`, when present. */
	codepage?: number;
	/** Workbook structure protection (`PROTECT` in the globals). */
	structureLocked: boolean;
	/** Workbook-level content this reader saw but does not model (VBA, pivot caches, ...). */
	unsupported: string[];
}

export type XlsReadErrorCode = 'encrypted' | 'unsupported-version' | 'corrupt';

/** Thrown by `readXlsWorkbook` for encrypted, pre-BIFF8 or unreadable input. */
export class XlsReadError extends Error {
	readonly code: XlsReadErrorCode;
	constructor(code: XlsReadErrorCode, message: string) {
		super(message);
		this.name = 'XlsReadError';
		this.code = code;
	}
}
