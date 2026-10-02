/**
 * One BIFF8 worksheet substream: cell records (`NUMBER`, `RK`, `MULRK`,
 * `LABEL`, `LABELSST`, `RSTRING`, `BOOLERR`, `BLANK`, `MULBLANK`, `FORMULA`
 * with its cached result and following `STRING`), shared (`SHRFMLA`) and
 * array (`ARRAY`) formula definitions, plus the layout, view, hyperlink and
 * comment records handled in `legacy-excel-workbook-sheet-extras.ts`.
 * Embedded chart substreams are skipped and reported as unsupported.
 *
 * @module legacy-excel-workbook-sheet
 */
import { decodeRk } from './legacy-excel-biff8.js';
import type { XlsBoundSheet, XlsGlobals } from './legacy-excel-workbook-globals.js';
import {
	formulaDefinition,
	formulaValue,
	resolveFormulas,
	type FormulaDefinition,
	type PendingFormula,
} from './legacy-excel-workbook-formulas.js';
import { BIFF_ERRORS } from './legacy-excel-workbook-ptg.js';
import { ByteReader, SegmentReader, type XlsRecord } from './legacy-excel-workbook-records.js';
import {
	applyPane,
	applyWindow2,
	defaultSheetView,
	parseColInfo,
	parseHyperlink,
	parseHyperlinkTooltip,
	parseMergedCells,
	parseNote,
	parseObj,
	parseRow,
	parseScl,
	parseSelection,
	parseTxoText,
} from './legacy-excel-workbook-sheet-extras.js';
import { paletteColor } from './legacy-excel-workbook-styles.js';
import type { XlsCell, XlsSheet } from './legacy-excel-workbook-types.js';

const UNSUPPORTED: Readonly<Record<number, string>> = {
	0x01b0: 'conditional formatting',
	0x0879: 'conditional formatting',
	0x01b2: 'data validation',
	0x009d: 'autofilter',
	0x00b0: 'pivot tables',
	0x0236: 'data tables',
	0x0868: 'protection ranges',
};

const OBJECT_KINDS: Readonly<Record<number, string>> = {
	0x05: 'charts',
	0x08: 'pictures',
};

function addUnique(list: string[], item: string): void {
	if (!list.includes(item)) list.push(item);
}

/** Parse the records of one sheet substream (`BOF` to its matching `EOF`). */
export function parseSheet(
	bound: XlsBoundSheet,
	records: XlsRecord[],
	start: number,
	globals: XlsGlobals,
): XlsSheet {
	const sheet: XlsSheet = {
		name: bound.name,
		state: bound.state,
		kind: bound.kind,
		cells: [],
		merges: [],
		columns: [],
		rows: [],
		view: defaultSheetView(),
		hyperlinks: [],
		comments: [],
		unsupported: [],
	};
	const formulas: PendingFormula[] = [];
	const shared = new Map<string, FormulaDefinition>();
	const arrays = new Map<string, FormulaDefinition>();
	const texts = new Map<number, string>();
	const notes: ReturnType<typeof parseNote>[] = [];
	const selections: ReturnType<typeof parseSelection>[] = [];
	let stringTarget: XlsCell | undefined;
	let lastObject: { type: number; id: number } | undefined;
	let frozen = false;
	let activePane = 3;
	let hasAutoFilter = false;
	let depth = 0;
	const cell = (r: ByteReader, value: XlsCell['value']): XlsCell => {
		const result: XlsCell = { row: r.u16(), col: r.u16(), xf: r.u16(), value };
		sheet.cells.push(result);
		return result;
	};

	for (let index = start; index < records.length; index++) {
		const record = records[index]!;
		const { opcode, data } = record;
		if (opcode === 0x0809) {
			depth++;
			if (depth > 1) addUnique(sheet.unsupported, 'charts');
			continue;
		}
		if (opcode === 0x000a) {
			depth--;
			if (depth <= 0) break;
			continue;
		}
		if (depth !== 1) continue;
		const r = new ByteReader(data);
		switch (opcode) {
			case 0x0203:
				cell(r, new ByteReader(data, 6).f64());
				break;
			case 0x027e:
				cell(r, decodeRk(new ByteReader(data, 6).u32()));
				break;
			case 0x00bd: {
				const row = r.u16();
				const first = r.u16();
				const count = Math.floor((data.length - 6) / 6);
				for (let i = 0; i < count; i++) {
					const xf = r.u16();
					sheet.cells.push({ row, col: first + i, xf, value: decodeRk(r.u32()) });
				}
				break;
			}
			case 0x0201:
				cell(r, null);
				break;
			case 0x00be: {
				const row = r.u16();
				const first = r.u16();
				const count = Math.floor((data.length - 6) / 2);
				for (let i = 0; i < count; i++)
					sheet.cells.push({ row, col: first + i, xf: r.u16(), value: null });
				break;
			}
			case 0x00fd: {
				const target = cell(r, '');
				target.value = globals.sst[r.u32()] ?? '';
				break;
			}
			case 0x0204:
			case 0x00d6: {
				const target = cell(r, '');
				const reader = SegmentReader.of({ ...record, data: data.subarray(6) });
				target.value = reader.xlString();
				break;
			}
			case 0x0205: {
				const target = cell(r, null);
				const raw = r.u8();
				target.value = r.u8() ? { error: BIFF_ERRORS[raw] ?? '#N/A' } : raw !== 0;
				break;
			}
			case 0x0006: {
				const value = formulaValue(data);
				const target = cell(r, value === 'string' ? '' : value);
				stringTarget = value === 'string' ? target : undefined;
				const cce = new ByteReader(data, 20).u16();
				formulas.push({
					cell: target,
					rgce: data.subarray(22, 22 + cce),
					extra: data.subarray(22 + cce),
				});
				break;
			}
			case 0x0207:
				if (stringTarget) stringTarget.value = SegmentReader.of(record).xlString();
				stringTarget = undefined;
				break;
			case 0x04bc: {
				const definition = formulaDefinition(data, 2);
				shared.set(`${definition.range.firstRow},${definition.range.firstCol}`, definition);
				break;
			}
			case 0x0221: {
				const definition = formulaDefinition(data, 6);
				arrays.set(`${definition.range.firstRow},${definition.range.firstCol}`, definition);
				break;
			}
			case 0x0200: {
				const firstRow = r.u32();
				const endRow = r.u32();
				const firstCol = r.u16();
				const endCol = r.u16();
				if (endRow > firstRow && endCol > firstCol)
					sheet.dimensions = { firstRow, lastRow: endRow - 1, firstCol, lastCol: endCol - 1 };
				break;
			}
			case 0x0208:
				sheet.rows.push(parseRow(data));
				break;
			case 0x007d:
				sheet.columns.push(parseColInfo(data));
				break;
			case 0x0055:
				sheet.defaultColWidth = r.u16();
				break;
			case 0x0099:
				sheet.standardWidth = r.u16() / 256;
				break;
			case 0x0225:
				r.skip(2);
				sheet.defaultRowHeight = r.u16() / 20;
				break;
			case 0x00e5:
				sheet.merges.push(...parseMergedCells(data));
				break;
			case 0x023e:
				frozen = applyWindow2(sheet.view, data);
				break;
			case 0x00a0: {
				const zoom = parseScl(data);
				if (zoom && zoom !== 100) sheet.view.zoom = zoom;
				break;
			}
			case 0x0041:
				activePane = applyPane(sheet.view, data, frozen);
				break;
			case 0x001d:
				selections.push(parseSelection(data));
				break;
			case 0x01b8: {
				const link = parseHyperlink(data);
				if (link) sheet.hyperlinks.push(link);
				break;
			}
			case 0x0800: {
				const tip = parseHyperlinkTooltip(data);
				const link = sheet.hyperlinks.find(
					(item) =>
						item.range.firstRow === tip.range.firstRow &&
						item.range.firstCol === tip.range.firstCol,
				);
				if (link && tip.text) link.tooltip = tip.text;
				break;
			}
			case 0x005d: {
				lastObject = parseObj(data);
				const kind = lastObject ? OBJECT_KINDS[lastObject.type] : undefined;
				if (kind) addUnique(sheet.unsupported, kind);
				else if (
					lastObject &&
					lastObject.type !== 0x19 &&
					!(lastObject.type === 0x14 && hasAutoFilter)
				)
					addUnique(sheet.unsupported, 'shapes and form controls');
				break;
			}
			case 0x01b6:
				if (lastObject?.type === 0x19) texts.set(lastObject.id, parseTxoText(record));
				break;
			case 0x001c:
				notes.push(parseNote(data));
				break;
			case 0x0862: {
				const icv = new ByteReader(data, 16).u32() & 0x7f;
				if (data.length >= 20 && icv !== 0x7f) sheet.tabColor = paletteColor(icv, globals.palette);
				break;
			}
			case 0x0012:
				if (r.u16()) sheet.protection ??= {};
				break;
			case 0x0013: {
				const hash = r.u16();
				if (hash)
					sheet.protection = { passwordHash: hash.toString(16).toUpperCase().padStart(4, '0') };
				break;
			}
			default: {
				if (opcode === 0x009d) hasAutoFilter = true;
				const label = UNSUPPORTED[opcode];
				if (label) addUnique(sheet.unsupported, label);
			}
		}
	}

	resolveFormulas(formulas, shared, arrays, globals);
	for (const note of notes) {
		sheet.comments.push({
			row: note.row,
			col: note.col,
			author: note.author,
			text: texts.get(note.id) ?? '',
			visible: note.visible,
		});
	}
	const selection = selections.find((item) => item.pane === activePane) ?? selections[0];
	if (selection) {
		sheet.view.activeCell = selection.active;
		if (selection.ranges.length) sheet.view.selection = selection.ranges;
	}
	return sheet;
}
