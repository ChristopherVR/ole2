/** A displayed cell value from a legacy BIFF8 worksheet preview. */
export interface OleSheetCell {
	value: string;
	isNumeric: boolean;
}

/** One bounded preview row. */
export interface OleSheetRow {
	cells: OleSheetCell[];
}

/** A read-only snapshot of one legacy workbook worksheet. */
export interface OleSheetGrid {
	sheetName: string;
	rows: OleSheetRow[];
}
