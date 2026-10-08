# DOC bookmark growth prerequisites

The current writer continues to refuse paragraph growth/shrink when standard bookmark tables are populated. No bookmark write support was added. This audit covers FIB pairs 21 (`SttbfBkmk`), 22 (`PlcfBkf`) and 23 (`PlcfBkl`).

## Primary layout facts

The FIB requires matching element counts across the names/start/end tables. Standard bookmark endpoints may reach the end of all document parts, so a main-story count alone is insufficient when other stories exist. Zero-length table pointers are undefined and must be ignored. [MS-DOC FibRgFcLcb97](https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-doc/0c9df81f-98d0-454e-ad84-b612cd05b1a4).

Names use UTF-16 extended strings, a two-byte count, `fExtend=0xffff` and `cbExtra=0`; names must be unique. The specification has tighter narrative name/count limits alongside looser field ceilings: a future implementation must reconcile both, rather than accepting the larger field limits alone. Preserve the original name bytes when relocating CPs. [MS-DOC SttbfBkmk](https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-doc/a65c98b7-8ca6-4220-9850-51703696f3ee).

For `n` bookmarks, the start PLC has `n+1` CPs followed by `n` four-byte FBKF records: `lcb=8n+4`. Each FBKF contains a two-byte `ibkl` and a two-byte BKC. `ibkl` is a unique zero-based end-table index; start order need not equal end order. [MS-DOC FBKF](https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-doc/4dfad7b0-37bb-443a-8933-3bb79d2c5994).

The end PLC has only `n+1` CPs: `lcb=4n+4`. It contains no reverse-index records. Each actual CP is the exclusive bookmark limit. [MS-DOC Plcfbkl](https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-doc/1c65c62e-2c88-4a85-8207-c7ec735fa7d8).

Both CP arrays are ordered; duplicate endpoints are permitted. Each paired start must not exceed its limit. The final sentinel is exactly one greater than the maximum permitted bookmark endpoint, not the field PLC's undefined sentinel. Bookmark ranges have special table-boundary rules. [MS-DOC Plcfbkf](https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-doc/82702014-0081-40af-8d4e-421db135f3ff).

A range containing any field marker must include the complete field; endpoints must share a document part. Column bookmarks additionally depend on table depth and cell/row boundaries encoded by BKC, which cannot be inferred from paragraph text alone. [MS-DOC Valid Selection](https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-doc/8d8fece5-bdbc-4258-8457-916540075e3f), [MS-DOC BKC](https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-doc/b59cbd01-2391-4b70-994c-b58104d3f4c0).

## Fixture audit

All seven checked-in DOC binaries were inspected with `readDocFib`/`readFcLcbAt`; every pair 21/22/23 has `lcb=0`. All report effective FIB version `0x0112` (274), including the extension rather than relying on the base FIB version. This audit therefore does not exercise populated bookmarks across every supported FIB version. SHA-256 values match `test/fixtures/provenance.json`.

| Fixture | Provenance | SHA-256 |
| --- | --- | --- |
| `ole-word-97.doc` | Existing historical fixture; origin/license unverified | `70f55e2046c0027174f6172755ef71ed2098978bba984d1e247d2365c33091d0` |
| `doc/header-field.doc` | Repository-owned Word-generated synthetic content | `fee2f2d258bf6c5949ae8456ca1d001cace2426a1a56edb7754acc23ec86dc26` |
| `doc/main-field.doc` | Repository-owned Word-generated synthetic content | `4df0889fa90aca6c272c61a19d68f133893a7efb451d062f1dfcdb12d1cb5e2e` |
| `doc/paragraph-alignment.doc` | Repository-owned Word-generated derivative | `a5ea2e6931374327fb8f5c4be50c241a9a742d9c5ede2143dd358ea1982e0ad1` |
| `doc/rich-runs.doc` | Repository-owned Word-generated synthetic content | `4b1cdf12f247fc235fdbee7232c3ed1d16d6380a98810b709b05e516cbfdd72f` |
| `doc/rich-size-runs.doc` | Repository-owned Word-generated derivative | `07acf328ddc3e6c72126df3402a5f184088d3bf56e46c0e54d9a4245401f1b21` |
| `doc/underline-runs.doc` | Repository-owned Word-generated derivative | `b788afb59be842da7baa0f65470ea298dd950851c7a0a411aa56c70046488154` |

The manifest names the six owned fixtures' PowerShell generators. None creates bookmark scenarios. The existing Word snapshot script does not capture bookmark names, endpoints or selected text. No new external binary was acquired.

## Admission decision

The primary layouts make a narrow implementation plausible: a main-story-only, non-column bookmark wholly outside the replacement could shift both sorted CP arrays and their exact sentinels, keeping the FBKF permutation, BKC/name bytes and original allocation. This is an implementation candidate, not verified behavior. Endpoints inside the replaced paragraph, collapsed bookmarks at replacement boundaries and bookmarks spanning the replacement require an explicit affinity policy; refuse them initially. A paragraph replacement includes its terminating mark, so visible-text boundaries are insufficient.

Before admission, validate table bounds/counts, index permutation, ordering/sentinels, legal selections and non-overlapping FIB allocation; retain all existing feature guards. Preserve formatting/opaque streams and prove transactional failure in clean and dirty models. Grow/shrink native Word-generated bookmark fixtures and independently check names, endpoint locations, selected text and save/reopen against an unchanged native-save control. Native Word is unavailable in this cloud environment. The format does not mandate a COM oracle; retaining refusal here is the project's evidence-based safety decision.

`test/ole-document-doc-bookmark-refusal.test.ts` builds one standard bookmark from these layouts in memory on the owned `rich-runs.doc` base. It tests refusal before/after the bookmark and clean/dirty model atomicity, plus table preservation during an allowed equal-length edit. It is neither a native-authored bookmark fixture nor proof that Word accepts the derived bytes.
