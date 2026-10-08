# MinkExcel

<img src="assets/icon.png" alt="MinkExcel white mink and XL icon on a solid Excel green squircle" width="160" height="160">

Small XLSX import and export for accounting workflows. This private
standalone TypeScript package has **zero npm dependencies**, including development
and peer dependencies. Its source uses standard JavaScript APIs and has no
Node/Bun imports, filesystem access, network access, or application framework imports.
It runs in browser bundles, Web Workers, Bun, and Node through a TypeScript bundler/runtime.

## Usage

```ts
import { Workbook, readWorkbook, writeWorkbook } from "minkexcel";

const book = new Workbook();
const sheet = book.addWorksheet("Customers");
sheet.addRow(["TIN", "Name", "Amount"]);
sheet.addRow(["001234567", "Ana & Co", 1234.56]);
sheet.getColumn(3).numFmt = "#,##0.00";

const bytes = await writeWorkbook(book);
const imported = await readWorkbook(bytes);
```

In the browser, obtain input with `new Uint8Array(await file.arrayBuffer())` and
create an output `Blob` from the returned bytes with MIME type
`application/vnd.openxmlformats-officedocument.spreadsheetml.sheet`.
For larger workbooks, run the package in a Web Worker to move XML processing off
the main thread. Read and write support an optional `AbortSignal`:

```ts
await readWorkbook(
  bytes,
  { rows: 10_000, columns: 100, cells: 100_000 },
  signal,
);
await writeWorkbook(book, signal);
```

## Supported scope

- Typed strings, numbers, booleans, dates, blanks, errors and cached formulas.
  Strings retain leading zeros, Unicode, whitespace and Excel escape sequences.
- Workbook metadata, multiple sheets, row gaps, widths, heights, fonts, solid
  fills, alignment, number formats, merged cells, frozen rows, filters, print
  titles/areas, page setup/margins and footers on export.
- Shared/inline strings, rich text flattened to text, hyperlink display text,
  cached formulas and 1900/1904 dates on import. Import reads values and uses
  number formats to identify dates; it does not preserve presentation for editing.
- Owned ZIP32 reader for STORE and DEFLATE, CRC32 checks, internal relationship
  resolution, XML parsing without DTDs/entities from external sources, size and
  nesting limits, and cancellation checkpoints during archive/row processing.

This is a focused implementation of common accounting workbook requirements.
It does not evaluate formulas, edit arbitrary workbooks without losing features,
read legacy `.xls`, support ZIP64/encryption, or implement charts, images, macros,
conditional formatting or the complete ExcelJS API.

Exports use ZIP DEFLATE compression through the standard `CompressionStream`
API. The package removes its zlib envelope to write the raw DEFLATE payload
required by ZIP; checksums and ZIP headers remain owned by the package. Export
requires a runtime that provides `CompressionStream('deflate')` and Web Streams.
See the [Compression Standard](https://compression.spec.whatwg.org/#supported-formats).
Import accepts both stored and compressed workbooks. The package builds workbook
XML and ZIP bytes in memory; it is not a streaming workbook writer.

Default import limits: 5 MiB input, 1,000 archive entries, 20 MiB per entry,
50 MiB total uncompressed data, compression ratio 200, 10,000 data rows plus the
header, 100 columns and a 100,000-cell rectangular extent per worksheet.
Callers can supply explicit limits. Inflate output is also bounded by the
entry's declared size and validated against its CRC32. External relationships
are never fetched.

## Tests and benchmark

Run from this project:

```sh
bun run test
```

Browser tests and the ExcelJS benchmark use optional development tools isolated
in `tools/package.json`, which has its own lockfile. The library manifest still
has no dependencies, devDependencies or peerDependencies. To run these tools:

```sh
bun install --cwd tools --ignore-scripts
bun run test:browser
bun run bench
```

Install Chromium with `bun tools/node_modules/@playwright/test/cli.js install chromium`
if it is not already available.

The unit tests and synthetic workbook fixtures live in `tests/`. They use
Bun's built-in test runner and Node's built-in compressor and inflater as independent
DEFLATE references; no external test packages are needed for the unit tests.
The browser harness in `tests/browser.ts` uses the optional Playwright tool
to test a browser bundle and a Web Worker.

The optional benchmark in `benchmarks/exceljs.ts` uses ExcelJS from the tools
project, not a library dependency. It compares equivalent
synthetic ledgers at 100, 1,000 and 10,000 rows, checks exported values outside
timed samples, and feeds the same compressed input to both readers. See
[measured results](benchmarks/RESULTS.md) and `benchmarks/results.json` for timing,
output size, runtime, hardware, sample count and scope differences.

ZIP layout follows the [PKWARE APPNOTE](https://pkware.cachefly.net/webdocs/casestudies/APPNOTE.TXT);
DEFLATE follows [RFC 1951](https://www.rfc-editor.org/rfc/rfc1951).
