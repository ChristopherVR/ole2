/** VSD's flag-byte / 4KiB sliding-window representation. Format facts checked
 * against Apache POI's HDGFLZW and libvisio's VSDInternalStream; independently
 * implemented here. No external decoder code is incorporated. */
export function decodeVsdBlock(input: Uint8Array, limit = 32 * 1024 * 1024): Uint8Array {
 if (!Number.isSafeInteger(limit) || limit < 0 || limit > 64 * 1024 * 1024) throw new Error('Invalid VSD decompression limit');
 const output = new Uint8Array(Math.min(limit, input.length * 9)), window = new Uint8Array(4096);
 let cursor = 0, position = 0;
 while (cursor < input.length) {
  const flags = input[cursor++]!;
  if (cursor === input.length) throw new Error('Truncated VSD compression flags');
  for (let bit = 0; bit < 8 && cursor < input.length; bit++) {
   if (flags & (1 << bit)) {
    if (position >= output.length) throw new Error('VSD decompression limit');
    const value = input[cursor++]!; window[position & 4095] = value; output[position++] = value;
   } else {
    if (cursor + 2 > input.length) throw new Error('Truncated VSD compression reference');
    const low = input[cursor++]!, high = input[cursor++]!;
    const length = (high & 15) + 3, address = (((high & 240) << 4) | low);
    const source = (address + 18) & 4095;
    if (length > output.length - position) throw new Error('VSD decompression limit');
    for (let i = 0; i < length; i++) {
     const value = window[(source + i) & 4095]!;
     window[position & 4095] = value; output[position++] = value;
    }
   }
  }
 }
 return output.slice(0, position);
}

/** Literal-only encoding deliberately trades compression for an unambiguous
 * preservation writer. It retains the decoded four-byte stream prefix. */
export function encodeVsdBlock(input: Uint8Array): Uint8Array {
 const output = new Uint8Array(input.length + Math.ceil(input.length / 8));
 let cursor = 0;
 for (let offset = 0; offset < input.length; offset += 8) {
  const count = Math.min(8, input.length - offset);
  output[cursor++] = (1 << count) - 1;
  output.set(input.subarray(offset, offset + count), cursor); cursor += count;
 }
 return output;
}
