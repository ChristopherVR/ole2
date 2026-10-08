/* Adapted from ChristopherVR/pptx-viewer packages/core/src/core/utils, Apache-2.0. Original source: https://github.com/ChristopherVR/pptx-viewer. */
/**
 * Types, constants, and error classes for the OLE2 parser.
 *
 * Reference: [MS-CFB] Compound Binary File Format
 * @see https://docs.microsoft.com/en-us/openspecs/windows_protocols/ms-cfb
 *
 * @module ole2-parser-types
 */

/* ------------------------------------------------------------------ */
/*  Constants                                                          */
/* ------------------------------------------------------------------ */

/** OLE2 Compound Binary File magic signature. */
export const OLE_MAGIC = new Uint8Array([0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1]);

/** Special sector index: end of chain. */
export const ENDOFCHAIN = 0xfffffffe;
/** Special sector index: free sector. */
export const FREESECT = 0xffffffff;
/** Special sector index: FAT sector. */
export const FATSECT = 0xfffffffd;
/** Special sector index: DIFAT sector. */
export const DIFSECT = 0xfffffffc;
/** Maximum regular sector index. */
export const MAXREGSECT = 0xfffffffa;

/** Directory entry object type: empty. */
export const ENTRY_TYPE_EMPTY = 0;
/** Directory entry object type: storage (folder). */
export const ENTRY_TYPE_STORAGE = 1;
/** Directory entry object type: stream (file). */
export const ENTRY_TYPE_STREAM = 2;
/** Sentinel for an absent directory link. */
export const NOSTREAM = 0xffffffff;
/** Directory entry object type: root entry. */
export const ENTRY_TYPE_ROOT = 5;

/** Directory entry size is always 128 bytes. */
export const DIR_ENTRY_SIZE = 128;

/* ------------------------------------------------------------------ */
/*  Types                                                              */
/* ------------------------------------------------------------------ */

/**
 * Parsed OLE2 directory entry.
 */
export interface Ole2DirectoryEntry {
	/** Entry name (UTF-16LE decoded). */
	name: string;
	/** Object type (ENTRY_TYPE_* constant). */
	type: number;
	/** Starting sector for this entry's data. */
	startSector: number;
	/** Size of the entry's data in bytes. */
	size: number;
	/** Index of the child directory entry (-1 if none). */
	childId: number;
	/** Index of the left sibling directory entry (-1 if none). */
	leftSiblingId: number;
	/** Index of the right sibling directory entry (-1 if none). */
	rightSiblingId: number;
	/**
	 * Storage CLSID (16 bytes), read from directory-entry offset 80. Present
	 * on every entry (all-zero when the host set no CLSID); callers that
	 * round-trip a container through `buildOle2` pass the root entry's value
	 * through as `rootClsid` so a re-serialized file keeps its original
	 * application identity (see `ole-sheet-xls-biff8.ts`'s CFB-rewrap helper).
	 */
	clsid: Uint8Array;
	/** Directory slot ID. `childId`/`leftSiblingId`/`rightSiblingId` refer to this, not to array positions. */
	id: number;
	/**
	 * Storage path below the root entry (`[]` for the root itself), resolved by
	 * walking the red-black sibling trees. Undefined for an entry no storage
	 * reaches (a stale slot some writers leave behind).
	 */
	path?: string[];
	/** Directory ID of the owning storage; -1 for the root and unreachable entries. */
	parentId: number;
	/** Red-black node color: 0 red, 1 black. */
	color: number;
	/** User-defined state bits (directory-entry offset 96). */
	stateBits: number;
	/** Creation time, when the entry records one. */
	created?: Date;
	/** Modification time, when the entry records one. */
	modified?: Date;
	/** Exact unsigned FILETIME ticks; preserves precision absent from Date. */
	createdFileTime?: bigint;
	/** Exact unsigned FILETIME ticks, including zero (unset). */
	modifiedFileTime?: bigint;
}

/** A non-fatal irregularity the reader accepted. No warning ever stands for fabricated stream data. */
export interface Ole2ParseWarning {
	code:
		| 'byte-order'
		| 'unaligned-length'
		| 'size-high-bits'
		| 'mini-fat-count'
		| 'chain-terminator'
		| 'directory-link'
		| 'unreachable-entry'
		| 'invalid-unreachable-entry';
	message: string;
}

/** Reader options. */
export interface Ole2ReadOptions {
	/**
	 * Reject every structural irregularity instead of accepting the safe ones
	 * (non-canonical byte-order mark, unaligned file length, header count
	 * drift, stale directory slots). Stream data is never truncated,
	 * zero-filled or otherwise invented in either mode.
	 */
	strict?: boolean;
}

/**
 * Parsed OLE2 compound file with stream access.
 */
export interface Ole2File {
	/** All non-empty directory entries, in directory slot order. */
	entries: Ole2DirectoryEntry[];
	/**
	 * Retrieve a named stream's binary data. A stream directly below the root
	 * wins over a same-named stream in a nested storage; the name compares
	 * exactly first and then case-insensitively ([MS-CFB] 2.6.4).
	 *
	 * @param name - The stream name to look up.
	 * @returns The stream data, or undefined if not found.
	 */
	getStream(name: string): Uint8Array | undefined;
	/**
	 * Retrieve a stream by storage path, as an array of names or a
	 * `/`-separated string. Names compare case-insensitively; a name without
	 * its leading control character (`SummaryInformation` for
	 * `\u0005SummaryInformation`) also matches when nothing matches exactly.
	 */
	getStreamByPath(path: string | readonly string[]): Uint8Array | undefined;
	/** Retrieve a stream by directory ID (see `Ole2DirectoryEntry.id`). */
	getStreamById(id: number): Uint8Array | undefined;
	/** Resolve a storage, stream or the root (`[]` or `'/'`) by path. */
	findEntry(path: string | readonly string[]): Ole2DirectoryEntry | undefined;
	/** Irregularities accepted while parsing. Empty for canonical files. */
	warnings: Ole2ParseWarning[];
}

/* ------------------------------------------------------------------ */
/*  Error                                                              */
/* ------------------------------------------------------------------ */

/**
 * Error thrown when OLE2 parsing fails.
 */
export class Ole2ParseError extends Error {
	public constructor(message: string) {
		super(message);
		this.name = 'Ole2ParseError';
	}
}
