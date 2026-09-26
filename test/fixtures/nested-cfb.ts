import { buildOle2 } from '../../src/ole2-parser-write.js';

/** Test-only CFB fixture builder with nested storage paths and duplicate leaf names. */
export function nestedCfb(streams: Array<{ path: string[]; bytes: Uint8Array }>): Uint8Array {
	const nodes = new Map<string, { path: string[]; bytes?: Uint8Array; id: number }>();
	for (const stream of streams)
		for (let depth = 1; depth <= stream.path.length; depth++) {
			const path = stream.path.slice(0, depth),
				key = path.join('/');
			if (!nodes.has(key)) nodes.set(key, { path, id: nodes.size + 1 });
			if (depth === stream.path.length) nodes.get(key)!.bytes = stream.bytes;
		}
	const flat = new Map(
		[...nodes.values()].map((node) => [
			`Slot${String(node.id).padStart(4, '0')}`,
			node.bytes ?? new Uint8Array(),
		]),
	);
	const result = new Uint8Array(buildOle2(flat));
	const view = new DataView(result.buffer);
	const sectorSize = 1 << view.getUint16(0x1e, true);
	// buildOle2 allocates directory sectors contiguously for these small fixtures.
	const directory = (view.getUint32(0x30, true) + 1) * sectorSize;
	const at = (id: number) => directory + id * 128;
	for (const node of nodes.values()) {
		const offset = at(node.id),
			name = node.path.at(-1)!;
		result.fill(0, offset, offset + 64);
		for (let i = 0; i < name.length; i++) view.setUint16(offset + 2 * i, name.charCodeAt(i), true);
		view.setUint16(offset + 64, (name.length + 1) * 2, true);
		result[offset + 66] = node.bytes ? 2 : 1;
		for (const field of [68, 72, 76]) view.setUint32(offset + field, 0xffffffff, true);
	}
	function attach(parent: number, path: string[]) {
		const children = [...nodes.values()].filter(
			(node) =>
				node.path.length === path.length + 1 && node.path.slice(0, -1).join('/') === path.join('/'),
		);
		children.sort(
			(a, b) =>
				a.path.at(-1)!.length - b.path.at(-1)!.length ||
				a.path.at(-1)!.toUpperCase().localeCompare(b.path.at(-1)!.toUpperCase()),
		);
		function tree(items: typeof children): number {
			if (!items.length) return 0xffffffff;
			const middle = Math.floor(items.length / 2),
				node = items[middle]!;
			view.setUint32(at(node.id) + 68, tree(items.slice(0, middle)), true);
			view.setUint32(at(node.id) + 72, tree(items.slice(middle + 1)), true);
			if (!node.bytes) attach(node.id, node.path);
			return node.id;
		}
		view.setUint32(at(parent) + 76, tree(children), true);
	}
	attach(0, []);
	return result;
}
