/**
 * Paragraph property lookup by file position for the structure readers.
 * Decodes every PAPX of a document once ([MS-DOC] 2.9.172, 2.9.173) and keeps
 * only what table detection needs, plus the table (sgc 5) sprms verbatim.
 *
 * @module doc-papx-index
 */
import { parseBteTable } from './ole-document-doc-fkp.js';
import { DocResourceLimitError } from './ole-document-doc-pieces.js';
import { lastSprm, operandI32, operandU8, readGrpprl, sprmGroup, type DocStructureSprm, type SprmBudget } from './doc-sprm-reader.js';

/** Direct paragraph properties relevant to table structure. */
export interface DocParaProps {
	readonly istd: number;
	readonly inTable: boolean;
	/** sprmPFTtp: this paragraph is a row end mark (outermost level). */
	readonly rowEnd: boolean;
	/** sprmPFInnerTableCell: nested-table cell mark. */
	readonly innerCell: boolean;
	/** sprmPFInnerTtp: nested-table row end mark. */
	readonly innerRowEnd: boolean;
	/** sprmPItap when present, else 1 inside a table and 0 outside. */
	readonly itap: number;
	/** Table sprms (TAP) carried by this paragraph, in order. */
	readonly tap: readonly DocStructureSprm[];
}

export interface DocPapxIndex {
	/** Properties of the paragraph whose mark is at byte `fc` of WordDocument, or undefined when uncovered. */
	at(fc: number): DocParaProps | undefined;
}

const NO_PROPS: DocParaProps = Object.freeze({ istd: 0, inTable: false, rowEnd: false, innerCell: false, innerRowEnd: false, itap: 0, tap: [] });

function decode(sprms: readonly DocStructureSprm[], istd: number): DocParaProps {
	const flag = (opcode: number) => (operandU8(lastSprm(sprms, opcode)) ?? 0) !== 0;
	const inTable = flag(0x2416);
	const itapRaw = operandI32(lastSprm(sprms, 0x6649));
	return {
		istd,
		inTable,
		rowEnd: flag(0x2417),
		innerCell: flag(0x244b),
		innerRowEnd: flag(0x244c),
		itap: itapRaw !== undefined && itapRaw >= 0 ? itapRaw : inTable ? 1 : 0,
		tap: sprms.filter((s) => sprmGroup(s.opcode) === 5),
	};
}

/** Build the index. `maxRuns` bounds the number of PAPX runs decoded. */
export function buildPapxIndex(
	word: Uint8Array,
	table: Uint8Array,
	papx: { fc: number; lcb: number },
	maxRuns: number,
	budget: SprmBudget,
): DocPapxIndex {
	if (papx.lcb < 4) return { at: () => undefined };
	if ((papx.lcb - 4) / 8 > maxRuns) throw new DocResourceLimitError('DOC formatting page count exceeds the processing limit');
	const bte = parseBteTable(table, papx);
	const view = new DataView(word.buffer, word.byteOffset, word.byteLength);
	const starts: number[] = [];
	const ends: number[] = [];
	const props: DocParaProps[] = [];
	for (const pnRaw of bte.pns) {
		const page = (pnRaw & 0x3fffff) * 512;
		if (page > word.length - 512) throw new Error('PAPX page outside WordDocument');
		const count = word[page + 511]!;
		const header = (count + 1) * 4 + count * 13;
		if (count === 0 || header > 511) throw new Error('Invalid PAPX page');
		if (starts.length + count > maxRuns) throw new DocResourceLimitError('DOC paragraph run count exceeds the processing limit');
		for (let i = 0; i < count; i++) {
			const start = view.getInt32(page + i * 4, true);
			const end = view.getInt32(page + (i + 1) * 4, true);
			if (end <= start || (starts.length > 0 && start < ends[ends.length - 1]!)) throw new Error('Invalid PAPX run range');
			let p = NO_PROPS;
			const offset = word[page + (count + 1) * 4 + i * 13]! * 2;
			if (offset !== 0) {
				if (offset < header || offset + 2 > 511) throw new Error('Invalid PAPX blob');
				const blob = page + offset;
				const cb = word[blob]!;
				const prefix = cb === 0 ? 2 : 1;
				const length = cb === 0 ? 2 + word[blob + 1]! * 2 : cb * 2;
				if (length < prefix + 2 || offset + length > 511) throw new Error('Invalid PAPX blob length');
				const istd = word[blob + prefix]! | (word[blob + prefix + 1]! << 8);
				p = decode(readGrpprl(word, blob + prefix + 2, blob + length, budget), istd);
			}
			starts.push(start);
			ends.push(end);
			props.push(p);
		}
	}
	return {
		at(fc: number) {
			let lo = 0;
			let hi = starts.length;
			while (lo < hi) {
				const mid = (lo + hi) >>> 1;
				if (ends[mid]! <= fc) lo = mid + 1;
				else hi = mid;
			}
			return lo < starts.length && fc >= starts[lo]! ? props[lo] : undefined;
		},
	};
}
