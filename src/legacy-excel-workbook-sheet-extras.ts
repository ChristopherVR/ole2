/**
 * Worksheet records around the cell table: column and row sizes, window
 * settings (`WINDOW2`, `SCL`, `PANE`, `SELECTION`), merged ranges,
 * hyperlinks (`HLINK`, `HLINKTOOLTIP`) and the text of legacy comments
 * (`OBJ` + `TXO` + `NOTE`).
 *
 * @module legacy-excel-workbook-sheet-extras
 */
import { ByteReader, SegmentReader, type XlsRecord } from './legacy-excel-workbook-records.js';
import type {
	XlsColumnInfo,
	XlsHyperlink,
	XlsRange,
	XlsRowInfo,
	XlsSheetView,
} from './legacy-excel-workbook-types.js';

/** `MERGEDCELLS`: a count and `Ref8` ranges. */
export function parseMergedCells(data: Uint8Array): XlsRange[] {
	const r = new ByteReader(data);
	const count = r.u16();
	const ranges: XlsRange[] = [];
	for (let i = 0; i < count && r.remaining >= 8; i++) ranges.push(readRef8(r));
	return ranges;
}

export function readRef8(r: ByteReader): XlsRange {
	return { firstRow: r.u16(), lastRow: r.u16(), firstCol: r.u16(), lastCol: r.u16() };
}

/** `COLINFO`: a column span's width (1/256 character units), format and flags. */
export function parseColInfo(data: Uint8Array): XlsColumnInfo {
	const r = new ByteReader(data);
	const firstCol = r.u16();
	const lastCol = r.u16();
	const width = r.u16() / 256;
	const xf = r.u16();
	const flags = r.u16();
	return {
		firstCol,
		lastCol: Math.min(lastCol, 255),
		width,
		xf,
		hidden: (flags & 0x01) !== 0,
		customWidth: (flags & 0x02) !== 0,
		bestFit: (flags & 0x04) !== 0,
		outlineLevel: (flags >> 8) & 0x07,
		collapsed: (flags & 0x1000) !== 0,
	};
}

/** `ROW`: height in twips, outline level and flags, optional default format. */
export function parseRow(data: Uint8Array): XlsRowInfo {
	const r = new ByteReader(data);
	const row = r.u16();
	r.skip(4);
	const height = (r.u16() & 0x7fff) / 20;
	r.skip(4);
	const flags = r.u16();
	const xf = r.u16() & 0x0fff;
	const info: XlsRowInfo = {
		row,
		height,
		customHeight: (flags & 0x40) !== 0,
		hidden: (flags & 0x20) !== 0,
		outlineLevel: flags & 0x07,
		collapsed: (flags & 0x10) !== 0,
	};
	if (flags & 0x80) info.xf = xf;
	return info;
}

/** A default sheet view (gridlines and headers on, 100% zoom). */
export function defaultSheetView(): XlsSheetView {
	return {
		showGridLines: true,
		showHeaders: true,
		showZeros: true,
		showFormulas: false,
		rightToLeft: false,
		selected: false,
		topRow: 0,
		leftCol: 0,
	};
}

/** `WINDOW2`: display flags, top-left cell and the normal-view zoom. */
export function applyWindow2(view: XlsSheetView, data: Uint8Array): boolean {
	const r = new ByteReader(data);
	const flags = r.u16();
	view.topRow = r.u16();
	view.leftCol = r.u16();
	view.showFormulas = (flags & 0x0001) !== 0;
	view.showGridLines = (flags & 0x0002) !== 0;
	view.showHeaders = (flags & 0x0004) !== 0;
	view.showZeros = (flags & 0x0010) !== 0;
	view.rightToLeft = (flags & 0x0040) !== 0;
	view.selected = (flags & 0x0200) !== 0;
	if (data.length >= 18) {
		r.skip(6);
		const zoom = r.u16();
		if (zoom && zoom !== 100) view.zoom = zoom;
	}
	return (flags & 0x0008) !== 0;
}

/** `SCL`: zoom as a fraction. */
export function parseScl(data: Uint8Array): number | undefined {
	const r = new ByteReader(data);
	const numerator = r.i16();
	const denominator = r.i16();
	return denominator > 0 ? Math.round((numerator / denominator) * 100) : undefined;
}

/** `PANE`: frozen rows/columns (or split twips) and the active pane. */
export function applyPane(view: XlsSheetView, data: Uint8Array, frozen: boolean): number {
	const r = new ByteReader(data);
	const x = r.u16();
	const y = r.u16();
	const topRow = r.u16();
	const leftCol = r.u16();
	const active = r.u8();
	if (frozen) {
		if (x || y) view.freeze = { rows: y, cols: x, topRow, leftCol };
	} else if (x || y) view.split = { x, y };
	return active;
}

/** `SELECTION`: the pane it belongs to, active cell and selected ranges. */
export function parseSelection(data: Uint8Array): {
	pane: number;
	active: { row: number; col: number };
	ranges: XlsRange[];
} {
	const r = new ByteReader(data);
	const pane = r.u8();
	const active = { row: r.u16(), col: r.u16() };
	r.skip(2);
	const count = r.u16();
	const ranges: XlsRange[] = [];
	for (let i = 0; i < count && r.remaining >= 6; i++)
		ranges.push({ firstRow: r.u16(), lastRow: r.u16(), firstCol: r.u8(), lastCol: r.u8() });
	return { pane, active, ranges };
}

const URL_MONIKER = 'e0c9ea79f9bace118c8200aa004ba90b';
const FILE_MONIKER = '0303000000000000c000000000000046';

function guid(r: ByteReader): string {
	let hex = '';
	for (let i = 0; i < 16; i++) hex += r.u8().toString(16).padStart(2, '0');
	return hex;
}

function hyperlinkString(r: ByteReader): string {
	return r.utf16(r.u32());
}

/** `HLINK`: the linked range, display text, URL or file moniker and in-workbook location. */
export function parseHyperlink(data: Uint8Array): XlsHyperlink | undefined {
	const r = new ByteReader(data);
	const range = readRef8(r);
	r.skip(16 + 4);
	const flags = r.u32();
	const link: XlsHyperlink = { range };
	if (flags & 0x10) link.display = hyperlinkString(r);
	if (flags & 0x80) hyperlinkString(r);
	if (flags & 0x01) {
		if (flags & 0x100) link.target = hyperlinkString(r);
		else {
			const clsid = guid(r);
			if (clsid === URL_MONIKER) {
				const size = r.u32();
				const end = r.pos + size;
				link.target = r.utf16(size / 2);
				r.pos = end;
			} else if (clsid === FILE_MONIKER) {
				const up = r.u16();
				const ansiLength = r.u32();
				let path = r.chars(ansiLength, false).replace(/\u0000+$/, '');
				r.skip(24);
				const unicodeSize = r.u32();
				if (unicodeSize > 0) {
					const bytes = r.u32();
					r.skip(2);
					path = r.chars(bytes / 2, true);
				}
				link.target = '../'.repeat(up) + path;
			} else return undefined;
		}
	}
	if (flags & 0x08) link.location = hyperlinkString(r);
	if (link.target === undefined && link.location === undefined) return undefined;
	return link;
}

/** `HLINKTOOLTIP`: the range it applies to and the tooltip text. */
export function parseHyperlinkTooltip(data: Uint8Array): { range: XlsRange; text: string } {
	const r = new ByteReader(data, 2);
	const range = readRef8(r);
	return { range, text: r.utf16((data.length - 10) / 2) };
}

/** `OBJ`: the common object data (`FtCmo`) type and id. */
export function parseObj(data: Uint8Array): { type: number; id: number } | undefined {
	const r = new ByteReader(data);
	if (r.u16() !== 0x15) return undefined;
	r.skip(2);
	return { type: r.u16(), id: r.u16() };
}

/** `TXO`: text box characters from the `CONTINUE` records that follow it (formatting runs skipped). */
export function parseTxoText(record: XlsRecord): string {
	const count = new ByteReader(record.data, 10).u16();
	if (count === 0 || record.continues.length === 0) return '';
	const reader = new SegmentReader(record.continues);
	const wide = (reader.u8() & 1) !== 0;
	return reader.continuedChars(count, wide).replace(/\r\n?/g, '\n');
}

/** `NOTE`: comment anchor, visibility, object id and author. */
export function parseNote(data: Uint8Array): {
	row: number;
	col: number;
	visible: boolean;
	id: number;
	author: string;
} {
	const r = new ByteReader(data);
	const row = r.u16();
	const col = r.u16();
	const flags = r.u16();
	const id = r.u16();
	return { row, col, visible: (flags & 0x02) !== 0, id, author: r.xlString() };
}
