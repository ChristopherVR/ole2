# CFB v4 preservation scope

`resizeCompoundFileStream` supports existing regular streams in v4 containers
with 4096-byte sectors and a header-only DIFAT (at most 109 FAT sectors).
Nested storage paths, duplicate leaf names in different storages and directory
entries in later sectors are supported. Directory-sector counts must match the
actual directory allocation chain. Stream lengths with a nonzero high DWORD
are refused rather than truncated.

Growth and shrinkage append replacement sectors and release the old allocation.
Old physical payload bytes are retained. Existing header padding, directory tree
pointers, CLSIDs, state bits, timestamps, unknown streams and slack remain byte
identical except target start/length and affected FAT/header allocation fields.
Same-length edits use the existing allocation-alias and encryption validator.
No input buffer is mutated; failure returns the original reference and reason.

Variable-length mini-stream edits and mini/regular transitions in v4, external
DIFAT expansion, new directory entries and stream lengths above the current
32-bit reader capacity remain refused. Allocations that overlap another stream
or metadata, directory ownership cycles and inconsistent directory counts fail.
This is container support: format callers must update their internal offsets
and lengths, and it does not establish native Office document fidelity.

`test/fixtures/v4-cfb.ts` generates neutral repository-owned fixtures from code,
with valid red-black storage trees and no Office application payloads. No new
third-party binary or personal document is included. Tests cover 4096-sector
boundaries, later directory sectors, FAT growth at the 1024-entry capacity,
repeated growth/shrinkage, input views with nonzero byte offsets, byte-exact
unrelated physical preservation and malformed/refused layouts. Native Office
validation is not claimed for these synthetic container tests.

Reference: [MS-CFB](https://learn.microsoft.com/en-us/openspecs/windows_protocols/ms-cfb/53989ce4-7b05-4f8d-829b-d08d6148375b),
sections 2.2 (header), 2.3 (directory), and 2.6 (FAT).
