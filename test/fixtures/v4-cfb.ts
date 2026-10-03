/** Repository-owned synthetic MS-CFB v4 fixture. No Office/personal payloads. */
export function v4Cfb(streams: Array<{ path: string[]; bytes: Uint8Array }>): Uint8Array {
	const END = 0xfffffffe, FREE = 0xffffffff, SECTOR = 4096;
	const nodes = new Map<string, { path: string[]; bytes?: Uint8Array; id: number; start: number }>();
	for (const stream of streams) {
		if (stream.bytes.length < 4096) throw new Error('Fixture only supports regular streams');
		for (let depth = 1; depth <= stream.path.length; depth++) {
			const path = stream.path.slice(0, depth), key = path.join('/');
			if (!nodes.has(key)) nodes.set(key, { path, id: nodes.size + 1, start: END });
			if (depth === stream.path.length) nodes.get(key)!.bytes = stream.bytes;
		}
	}
	const directories = Math.ceil((nodes.size + 1) / 32);
	let sectors = directories;
	for (const node of nodes.values()) if (node.bytes) {
		node.start = sectors;
		sectors += Math.ceil(node.bytes.length / SECTOR);
	}
	let fats = 1;
	while (sectors + fats > fats * 1024) fats++;
	if (fats > 109) throw new Error('Fixture DIFAT is unsupported');
	const out = new Uint8Array((sectors + fats + 1) * SECTOR), view = new DataView(out.buffer);
	out.set([0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1]);
	view.setUint16(0x18, 0x3e, true); view.setUint16(0x1a, 4, true); view.setUint16(0x1c, 0xfffe, true);
	view.setUint16(0x1e, 12, true); view.setUint16(0x20, 6, true);
	view.setUint32(0x28, directories, true); view.setUint32(0x2c, fats, true);
	view.setUint32(0x30, 0, true); view.setUint32(0x38, 4096, true);
	view.setUint32(0x3c, END, true); view.setUint32(0x44, END, true);
	for (let i = 0; i < 109; i++) view.setUint32(0x4c + i * 4, i < fats ? sectors + i : FREE, true);
	for (let i = 0; i < fats; i++) out.fill(0xff, (sectors + i + 1) * SECTOR, (sectors + i + 2) * SECTOR);
	const fat = (id: number, next: number) => view.setUint32((sectors + Math.floor(id / 1024) + 1) * SECTOR + id % 1024 * 4, next, true);
	for (let id = 0; id < directories; id++) fat(id, id + 1 === directories ? END : id + 1);
	for (let i = 0; i < fats; i++) fat(sectors + i, 0xfffffffd);
	const at = (id: number) => SECTOR + id * 128;
	const entry = (id: number, name: string, type: number, start: number, size: number) => {
		const offset = at(id);
		for (let i = 0; i < name.length; i++) view.setUint16(offset + i * 2, name.charCodeAt(i), true);
		view.setUint16(offset + 64, (name.length + 1) * 2, true);
		out[offset + 66] = type; out[offset + 67] = 1;
		for (const field of [68, 72, 76]) view.setUint32(offset + field, FREE, true);
		view.setUint32(offset + 116, start, true); view.setUint32(offset + 120, size, true);
	};
	entry(0, 'Root Entry', 5, END, 0);
	for (const node of nodes.values()) {
		entry(node.id, node.path.at(-1)!, node.bytes ? 2 : 1, node.start, node.bytes?.length ?? 0);
		if (node.bytes) {
			out.set(node.bytes, (node.start + 1) * SECTOR);
			const count = Math.ceil(node.bytes.length / SECTOR);
			for (let i = 0; i < count; i++) fat(node.start + i, i + 1 === count ? END : node.start + i + 1);
		}
	}
	const attach = (parent: number, path: string[]) => {
		const children = [...nodes.values()].filter(n => n.path.length === path.length + 1 && n.path.slice(0, -1).join('/') === path.join('/'));
		children.sort((a, b) => a.path.at(-1)!.length - b.path.at(-1)!.length || a.path.at(-1)!.toUpperCase().localeCompare(b.path.at(-1)!.toUpperCase()));
		// Color the incomplete deepest level red so every root-to-leaf path
		// has the same black height (as required by MS-CFB directory trees).
		const redLevel = Math.floor(Math.log2(children.length + 1));
		const tree = (items: typeof children, depth = 0): number => {
			if (!items.length) return FREE;
			const middle = Math.floor(items.length / 2), n = items[middle]!;
			out[at(n.id) + 67] = depth === redLevel ? 0 : 1;
			view.setUint32(at(n.id) + 68, tree(items.slice(0, middle), depth + 1), true);
			view.setUint32(at(n.id) + 72, tree(items.slice(middle + 1), depth + 1), true);
			if (!n.bytes) attach(n.id, n.path);
			return n.id;
		};
		view.setUint32(at(parent) + 76, tree(children), true);
	};
	attach(0, []);
	return out;
}
