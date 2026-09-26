import { rm, lstat } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { dirname, join } from 'node:path';
import assert from 'node:assert/strict';
const root = fileURLToPath(new URL('../', import.meta.url));
const target = fileURLToPath(new URL('../dist', import.meta.url));
assert.equal(target, join(root, 'dist'), 'Refuse cleanup outside package dist');
const stat = await lstat(target).catch((error) => {
	if (error.code !== 'ENOENT') throw error;
});
assert(!stat?.isSymbolicLink(), 'Refuse cleanup of linked output directory');
await rm(target, { recursive: true, force: true });
