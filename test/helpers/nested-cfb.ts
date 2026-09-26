/** Convert one root stream into a storage and move its stream entry underneath it. */
export function nestRootStream(
	bytes: Uint8Array,
	streamName: string,
	storageName: string,
): Uint8Array {
	const out = bytes.slice();
	const view = new DataView(out.buffer);
	const dirSector = view.getUint32(0x30, true);
	const dirOffset = (dirSector + 1) * (1 << view.getUint16(0x1e, true));
	const sectorSize = 1 << view.getUint16(0x1e, true);
	let streamId = -1,
		emptyId = -1;
	for (let id = 0; id < sectorSize / 128; id++) {
		const at = dirOffset + id * 128;
		if (out[at + 66] === 0 && emptyId < 0) emptyId = id;
		if (out[at + 66] === 2) {
			const length = view.getUint16(at + 64, true);
			let name = '';
			for (let i = 0; i < length - 2; i += 2)
				name += String.fromCharCode(view.getUint16(at + i, true));
			if (name.toLowerCase() === streamName.toLowerCase()) streamId = id;
		}
	}
	if (streamId < 0 || emptyId < 0)
		throw new Error('Fixture needs a root stream and an empty directory slot');
	const streamAt = dirOffset + streamId * 128;
	const childAt = dirOffset + emptyId * 128;
	const childEntry = out.slice(streamAt, streamAt + 128);
	const setName = (at: number, name: string) => {
		for (let i = 0; i < 32; i++) view.setUint16(at + i * 2, 0, true);
		for (let i = 0; i < name.length; i++) view.setUint16(at + i * 2, name.charCodeAt(i), true);
		view.setUint16(at + 64, (name.length + 1) * 2, true);
	};
	out.set(childEntry, childAt);
	setName(childAt, streamName);
	view.setUint32(childAt + 68, 0xffffffff, true);
	view.setUint32(childAt + 72, 0xffffffff, true);
	view.setUint32(childAt + 76, 0xffffffff, true);
	setName(streamAt, storageName);
	out[streamAt + 66] = 1;
	view.setUint32(streamAt + 76, emptyId, true);
	view.setUint32(streamAt + 116, 0xfffffffe, true);
	view.setUint32(streamAt + 120, 0, true);
	view.setUint32(streamAt + 124, 0, true);
	return out;
}
