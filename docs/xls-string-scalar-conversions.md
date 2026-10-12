# XLS string scalar conversions

Existing BIFF8 `LABELSST` cells can become finite numbers, booleans or defined
Excel errors through the existing `XlsCellNode.value` setter:

```ts
const document = parseXls(bytes);
document.sheets[0].cell(0, 0).value = 37.125;
document.sheets[0].cell(1, 0).value = true;
document.sheets[0].cell(4, 0).value = { error: '#N/A' };
const saved = document.serialize();
```

Coordinates above refer to the repository-owned rich-string fixture. The same
checks apply to any eligible cell, with no automatic parsing of numeric text.
`editXlsStringScalarCell` accepts CFB or bare Workbook bytes;
`editXlsStringScalarWorkbookStream` accepts a bare BIFF8 stream. These lower-level
APIs return `edited` or `unchanged`; refusal returns the original input object.

## Preservation and refusals

The writer validates all LABELSST references and the serialized SST reference
count before replacing a record. It reduces that count by one and retains every
existing SST entry and CONTINUE byte, including unused rich/phonetic entries.
Shared aliases keep their original string and formatting. The selected cell's
XF survives; rich formatting attached to its former text does not apply to a
new numeric/boolean/error value. Known BOUNDSHEET/INDEX/DBCELL offsets are
relocated through the existing checked replacement path. Intermediate NUMBER
records and altered counters are never exposed if conversion fails.

Formula caches, legacy LABEL records, merged followers, malformed SST counts or
indices, unsupported pointer records and unsafe container allocation refuse.
Null clearing, nonfinite numbers and undefined error codes remain unsupported.
Rejected setters retain bytes, revision and dirty state. Existing cell handles
continue to read the new value, and formula caches are retained with
`recalculationRequired` set rather than evaluated.

## Native evidence

On 12 October 2026, the production writer edited
`test/fixtures/xls/workbook-rich-strings.xls` in five places: a plain shared string
to 37.125, rich text to true, a continued Unicode string to 8.5, text to #N/A, and
a later-sheet string to false. Microsoft Excel 16.0 opened the output through the
repository's bounded COM tool with normal corrupt-load mode, saved a copy and
reopened it. The declared values/error display, number formats and merge areas
passed; every unrelated field captured by the snapshot tool matched the original.

This is evidence for the captured fields in that fixture, not complete rendering,
all rich-text properties or universal native fidelity. The v4 and Buffer cases
have codec/pointer regression coverage, but this native run used the v3 fixture.

After building, generate files into a dedicated output directory:

```sh
node scripts/prepare-native-xls-string-scalar-edits.mjs OUTPUT_DIRECTORY
```

Run `scripts/native-office-bounded.ps1` on `original.xls` and `edited.xls`, both
with `-CaptureCellTypes`, writing `original.json` and `edited.json`. For the
edited input, set `-SaveCopyPath` to `native-saved.xls`; capture that saved file
as `reopened.json` with the same switch. Then verify declared edits and unrelated
fields:

```sh
node scripts/verify-native-xls-string-scalar-edits.mjs OUTPUT_DIRECTORY
```

The native tools disable macros and use explicit fixture paths, isolated owned
Office processes and bounded execution. They are developer checks, not part of
the cross-platform runtime or the default test suite.

The recorded native run used these SHA-256 digests (native save output can vary
between application versions/runs):

| File | SHA-256 |
| --- | --- |
| Original fixture | `8a6ef91b5e77c8334f40b5fe6b09f81d0e38aa98ba93deb464e93cd2dee2cf87` |
| Production writer output | `0f033e8f2e9a34725c9907ead27ae33e9349546d5747fbd234c14a479299ae94` |
| Native saved copy | `d89dc252baa71c37ee7cc311bb968da6aed203eb8a0ce73dcf25bce7b7e338b0` |
