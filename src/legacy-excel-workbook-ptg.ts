/**
 * BIFF8 parsed-formula (`Rgce`) decoding to A1 formula text.
 *
 * Covers the operand, operator, control and function tokens Excel writes
 * for ordinary worksheet formulas, defined names and shared formulas:
 * constants (`PtgInt`, `PtgNum`, `PtgStr`, `PtgBool`, `PtgErr`,
 * `PtgMissArg`, `PtgArray`), references (`PtgRef`, `PtgArea`, `PtgRefN`,
 * `PtgAreaN`, `PtgRef3d`, `PtgArea3d` and their `#REF!` forms), names
 * (`PtgName`, `PtgNameX`), all binary/unary operators, `PtgParen`,
 * `PtgAttr` (sum, if, choose, goto, volatile, whitespace) and functions
 * (`PtgFunc`, `PtgFuncVar`, including add-in calls through id 255). The
 * `PtgMem*` tokens only cache reference evaluation and are skipped. Any other
 * token makes the whole decode return `undefined` rather than a guess.
 *
 * @module legacy-excel-workbook-ptg
 */
import { XLS_FUNCTIONS, XLS_USER_FUNCTION_ID } from './legacy-excel-workbook-functions.js';
import {
	areaText,
	BIFF_ERRORS,
	BINARY,
	cellText,
	formatNumber,
	readArrayConstant,
	readLoc,
	stringLiteral,
	type PtgContext,
	type PtgOptions,
} from './legacy-excel-workbook-ptg-text.js';
import { ByteReader } from './legacy-excel-workbook-records.js';

export {
	BIFF_ERRORS,
	columnName,
	quoteSheetName,
	type PtgContext,
	type PtgOptions,
} from './legacy-excel-workbook-ptg-text.js';

/**
 * Decode a parsed formula to text without the leading `=`. `rgce` holds the
 * tokens, `extra` the trailing data (`PtgArray` constants, `PtgMemArea`
 * reference caches). Returns `undefined` for any token it cannot represent.
 */
export function decodeFormula(
	rgce: Uint8Array,
	extra: Uint8Array,
	context: PtgContext,
	options: PtgOptions,
): string | undefined {
	try {
		return decode(rgce, new ByteReader(extra), context, options);
	} catch {
		return undefined;
	}
}

class Unsupported extends Error {}

function decode(
	rgce: Uint8Array,
	extra: ByteReader,
	context: PtgContext,
	options: PtgOptions,
): string | undefined {
	const reader = new ByteReader(rgce);
	const stack: string[] = [];
	let space = '';
	let closeSpace = '';
	const pop = (): string => {
		const value = stack.pop();
		if (value === undefined) throw new Unsupported();
		return value;
	};
	const push = (text: string): void => {
		stack.push(space + text);
		space = '';
	};
	const call = (name: string, argc: number): void => {
		const args: string[] = [];
		for (let i = 0; i < argc; i++) args.unshift(pop());
		let fn = name;
		if (name === '') {
			const first = args.shift();
			if (first === undefined) throw new Unsupported();
			fn = first.trim().replace(/^_xlfn\./, '');
		}
		push(`${fn}(${args.join(',')}${closeSpace})`);
		closeSpace = '';
	};
	const prefix = (ixti: number): string => {
		const sheet = context.sheetPrefix(ixti);
		if (sheet === undefined) throw new Unsupported();
		return `${sheet}!`;
	};
	while (reader.remaining > 0) {
		const ptg = reader.u8();
		if (ptg >= 0x03 && ptg <= 0x11) {
			const right = pop();
			const left = pop();
			const symbol = BINARY[ptg] ?? '';
			stack.push(`${left}${space}${symbol}${right}`);
			space = '';
			continue;
		}
		switch (ptg) {
			case 0x12:
				push(`+${pop()}`);
				continue;
			case 0x13:
				push(`-${pop()}`);
				continue;
			case 0x14:
				stack.push(`${pop()}${space}%`);
				space = '';
				continue;
			case 0x15:
				push(`(${pop()}${closeSpace})`);
				closeSpace = '';
				continue;
			case 0x16:
				push('');
				continue;
			case 0x17:
				push(stringLiteral(reader.shortXlString()));
				continue;
			case 0x19:
				readAttr(reader, stack, (kind, text) => {
					if (kind === 'open') space += text;
					else if (kind === 'close') closeSpace += text;
				});
				continue;
			case 0x1c:
				push(BIFF_ERRORS[reader.u8()] ?? '#N/A');
				continue;
			case 0x1d:
				push(reader.u8() ? 'TRUE' : 'FALSE');
				continue;
			case 0x1e:
				push(String(reader.u16()));
				continue;
			case 0x1f:
				push(formatNumber(reader.f64()));
				continue;
		}
		if (ptg < 0x20 || ptg > 0x7f) return undefined;
		const base = ptg & 0x1f;
		switch (base) {
			case 0x00:
				reader.skip(7);
				push(readArrayConstant(extra));
				break;
			case 0x01: {
				const info = XLS_FUNCTIONS.get(reader.u16());
				if (!info || info.argc < 0) return undefined;
				call(info.name, info.argc);
				break;
			}
			case 0x02: {
				const argc = reader.u8() & 0x7f;
				const tab = reader.u16();
				if (tab & 0x8000) return undefined;
				if (tab === XLS_USER_FUNCTION_ID) call('', argc);
				else {
					const info = XLS_FUNCTIONS.get(tab);
					if (!info) return undefined;
					call(info.name, argc);
				}
				break;
			}
			case 0x03: {
				const name = context.name(reader.u32());
				if (name === undefined) return undefined;
				push(name);
				break;
			}
			case 0x04:
			case 0x0c: {
				const row = reader.u16();
				push(cellText(readLoc(row, reader.u16(), base === 0x0c, options)));
				break;
			}
			case 0x05:
			case 0x0d: {
				const rows = [reader.u16(), reader.u16()] as const;
				const cols = [reader.u16(), reader.u16()] as const;
				const relative = base === 0x0d;
				push(
					areaText(
						readLoc(rows[0], cols[0], relative, options),
						readLoc(rows[1], cols[1], relative, options),
					),
				);
				break;
			}
			case 0x06: // PtgMemArea: skip header and its cached reference list in the extra data
				reader.skip(6);
				extra.skip(extra.u16() * 8);
				break;
			case 0x07:
			case 0x08:
				reader.skip(6);
				break;
			case 0x09:
				reader.skip(2);
				break;
			case 0x0a:
				reader.skip(4);
				push('#REF!');
				break;
			case 0x0b:
				reader.skip(8);
				push('#REF!');
				break;
			case 0x19: {
				const ixti = reader.u16();
				const index = reader.u16();
				reader.skip(2);
				const name = context.externName(ixti, index);
				if (name === undefined) return undefined;
				push(name);
				break;
			}
			case 0x1a: {
				const sheet = prefix(reader.u16());
				const row = reader.u16();
				push(sheet + cellText(readLoc(row, reader.u16(), options.shared === true, options)));
				break;
			}
			case 0x1b: {
				const sheet = prefix(reader.u16());
				const rows = [reader.u16(), reader.u16()] as const;
				const cols = [reader.u16(), reader.u16()] as const;
				const relative = options.shared === true;
				push(
					sheet +
						areaText(
							readLoc(rows[0], cols[0], relative, options),
							readLoc(rows[1], cols[1], relative, options),
						),
				);
				break;
			}
			case 0x1c:
				push(`${prefix(reader.u16())}#REF!`);
				reader.skip(4);
				break;
			case 0x1d:
				push(`${prefix(reader.u16())}#REF!`);
				reader.skip(8);
				break;
			default:
				return undefined;
		}
	}
	return stack.length === 1 ? stack[0] : undefined;
}

/** `PtgAttr`: control tokens that do not change the formula text, `SUM` shortcut, whitespace. */
function readAttr(
	reader: ByteReader,
	stack: string[],
	whitespace: (kind: 'open' | 'close', text: string) => void,
): void {
	const flags = reader.u8();
	if (flags & 0x04) {
		const count = reader.u16();
		reader.skip((count + 1) * 2);
		return;
	}
	if (flags & 0x40) {
		const type = reader.u8();
		const count = reader.u8();
		const char = type % 2 === 1 && type < 6 ? '\n' : ' ';
		if (type === 0 || type === 1 || type === 2 || type === 3)
			whitespace('open', char.repeat(count));
		else if (type === 4 || type === 5) whitespace('close', char.repeat(count));
		return;
	}
	reader.skip(2);
	if (flags & 0x10) {
		const value = stack.pop();
		if (value === undefined) throw new Unsupported();
		stack.push(`SUM(${value})`);
	}
}
