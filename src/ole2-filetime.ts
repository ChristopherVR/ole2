/** Lossless MS-CFB FILETIME conversion (100 ns ticks since 1601-01-01). */
const EPOCH = 116444736000000000n;
const MAX = 0xffffffffffffffffn;

export function fileTimeDate(ticks: bigint): Date | undefined {
	if (ticks === 0n) return undefined;
	const delta = ticks - EPOCH;
	// BigInt division truncates toward zero; dates before 1970 need floor.
	const ms = delta >= 0n ? delta / 10000n : (delta - 9999n) / 10000n;
	const date = new Date(Number(ms));
	return Number.isNaN(date.getTime()) ? undefined : date;
}

/** Keep sub-millisecond ticks unless the caller changed the Date view. */
export function writeFileTime(view: DataView, offset: number, date?: Date, raw?: bigint): void {
	if (raw !== undefined && (typeof raw !== 'bigint' || raw < 0n || raw > MAX)) {
		throw new Error('CFB FILETIME must be an unsigned 64-bit bigint');
	}
	const originalDate = raw === undefined ? undefined : fileTimeDate(raw);
	if (raw !== undefined && (date === undefined || date.getTime() === originalDate?.getTime())) {
		view.setBigUint64(offset, raw, true);
		return;
	}
	if (date === undefined) return;
	if (!Number.isFinite(date.getTime())) throw new Error('Invalid CFB timestamp');
	const ticks = BigInt(date.getTime()) * 10000n + EPOCH;
	if (ticks < 0n || ticks > MAX) throw new Error('CFB timestamp is outside FILETIME range');
	view.setBigUint64(offset, ticks, true);
}
