# MinkExcel

<img src="assets/icon.png" alt="MinkExcel white mink and XL icon on a solid Excel green squircle" width="160" height="160">

[![npm version](https://img.shields.io/npm/v/minkexcel?style=plastic&color=orange)](https://www.npmjs.com/package/minkexcel)
[![npm monthly downloads](https://img.shields.io/npm/dm/minkexcel?style=plastic&label=downloads&color=brightgreen)](https://www.npmjs.com/package/minkexcel)
[![Documentation](https://img.shields.io/badge/docs-MinkExcel-217346?style=plastic)](https://accntech.github.io/minkexcel/)
[![GitHub stars](https://img.shields.io/github/stars/accntech/minkexcel?style=social&label=Stars)](https://github.com/accntech/minkexcel/stargazers)

A simple, fast replacement for common ExcelJS XLSX import/export
workflows, with **zero runtime dependencies**. Export application data, create
formatted reports and read spreadsheet values in browsers, Web Workers, Bun
and Node.js. The package ships ES modules and TypeScript declarations.

MinkExcel prioritizes fast processing and dependable file handling: bounded
imports, CRC32 validation, strict XML checks and preservation of literal cell
text. See the [ExcelJS comparison](#minkexcel-vs-exceljs) for measured speed and
file-integrity results.

Import extracts values by default. Use `preserveTemplate: true` to update cell
values in an existing workbook while retaining its original formatting, layout,
protection, comments and other archive parts. Template mode retains those parts
without exposing a complete presentation editing API.

## Documentation site

Read the [MinkExcel documentation](https://accntech.github.io/minkexcel/).

The static documentation site includes guides, the public API, an ExcelJS
comparison and interactive graphs of the recorded benchmarks. To build and
preview it locally:

```sh
npm run docs:build
npm run docs:preview
```

See [the site guide](docs/README.md) for GitHub Pages publishing and editing.

## Installation and compatibility

```sh
npm install minkexcel
```

The package is ESM-only; use `import`. Export requires
`CompressionStream("deflate")` and Web Streams. Import uses the package's own
inflater and does not require `DecompressionStream`. The library performs no
filesystem or network I/O; your application supplies and saves the bytes.

Verified on October 8, 2026:

| Environment | Verified version | Coverage |
| --- | --- | --- |
| Bun | 1.4.0 | README examples, unit tests, interoperability and benchmarks |
| Node.js | 24.12.0 | README examples and benchmarks using the built ES modules |
| Chromium | 153.0.8010.12 | Browser bundle and module Web Worker import/export |

These are verified versions, not minimum supported versions. Other browsers
and runtime versions need validation in your application.

The full browser ESM bundle measures **29.9 KiB minified** (30,624 bytes) and
**11.0 KiB gzipped** (11,301 bytes), using Bun 1.4.0 on October 8, 2026. This
covers all public exports and excludes declarations and the icon. See
[contributor instructions](https://github.com/accntech/minkexcel/blob/main/CONTRIBUTING.md)
to reproduce the measurement.

### Tree-shaking

Use named imports in an application build with tree-shaking enabled:

```ts
import { Workbook, writeWorkbook } from "minkexcel";
```

MinkExcel preserves ES modules, declares `sideEffects: false` and initializes
fixed decompression tables only when needed. An export-only bundle can exclude
the reader and inflater; a model-only bundle can exclude all ZIP and XML code.

Consumer bundles measured with Bun 1.4.0 on October 8, 2026:

| Public imports | Minified bytes | Gzip bytes |
| --- | ---: | ---: |
| `Workbook` | 4,532 | 1,763 |
| `Workbook`, `writeWorkbook` | 16,630 | 6,090 |
| `readWorkbook` | 19,150 | 7,513 |
| All public exports | 30,624 | 11,301 |

Sizes depend on the bundler and the APIs your application uses. Methods on
retained model classes generally remain. Direct browser or Node.js imports
without a bundling step do not remove unused code or installed package files.

## Quick start

### Export a workbook

```ts
import { Workbook, writeWorkbook } from "minkexcel";

const book = new Workbook();
const sheet = book.addWorksheet("Products");
sheet.addRow(["SKU", "Name", "Price"]);
sheet.addRow(["001234567", "Desk lamp", 1234.56]);
sheet.getColumn(3).numFmt = "#,##0.00";

const bytes = await writeWorkbook(book); // Uint8Array containing an XLSX file
```

Keep identifiers as strings to retain leading zeros. Number formats control
Excel's display; they do not change the stored value.

### Import and read values

Using `bytes` from the export above:

```ts
import { readWorkbook } from "minkexcel";

const imported = await readWorkbook(bytes);
const products = imported.getWorksheet("Products");
if (!products) throw new Error("Products worksheet is missing.");

console.log(products.getCell("A2").value); // "001234567"
console.log(products.getCell(2, 3).value); // 1234.56

products.eachRow((row, rowNumber) => {
  if (rowNumber === 1) return; // Skip this workbook's header.
  console.log(row.getCell(1).value, row.getCell(3).value);
});
```

Use `imported.worksheets` to access all worksheets. `getWorksheet(name)` returns
`undefined` when there is no exact name match. `eachRow` and `eachCell` visit
populated entries in position order, skipping blank rows and cells.

### Browser files and downloads

```ts
import { readWorkbook, writeWorkbook, type Workbook } from "minkexcel";

async function importFile(file: File): Promise<Workbook> {
  return readWorkbook(new Uint8Array(await file.arrayBuffer()));
}

async function downloadWorkbook(book: Workbook): Promise<void> {
  const bytes = await writeWorkbook(book);
  const blob = new Blob([bytes.slice().buffer], {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  });
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = "products.xlsx";
  document.body.append(link);
  link.click();
  link.remove();
  setTimeout(() => URL.revokeObjectURL(url), 1_000);
}
```

For larger workbooks, run import and export in a module Web Worker to move XML
processing off the main thread. Workbook XML and ZIP bytes are built in memory;
there is no streaming workbook writer.

## API conventions

Row and column numbers are **one-based**: `getCell(2, 3)` refers to `C2`.
Cell addresses use uppercase letters, such as `"C2"`.

Arrays supplied to `addRow` or assigned to `row.values` are **zero-based**,
but reading `row.values` produces a **one-based sparse array**:

```ts
const row = sheet.addRow(["001234567", "Desk lamp", 1234.56]);
console.log(row.values[0]); // undefined
console.log(row.values[1]); // "001234567"
console.log(row.values[3]); // 1234.56
const values = row.values.slice(1); // ["001234567", "Desk lamp", 1234.56]
```

Blank cells create gaps in the returned array; `slice(1)` retains those gaps.
Prefer `getCell(column).value` when you need a specific field. `getRow` and
`getCell` create missing rows or cells; a newly created cell has value `null`.

`sheet.rowCount`, `sheet.columnCount` and `row.cellCount` describe the furthest
created positions, not the number of populated entries. Column definitions
alone do not increase `sheet.columnCount`.

`cell.value` returns a string, number, boolean, `Date`, `null`, formula object
or error object. `cell.text` returns a plain string representation or a formula's
cached result; it does not apply Excel number formatting.

## Formatted report recipe

This example combines a merged title, frozen headings, dates, numeric
number formatting, a cached total, filters and print settings:

```ts
import { Workbook, writeWorkbook } from "minkexcel";

const report = new Workbook();
report.creator = "Product team";
const inventory = report.addWorksheet("Inventory", {
  views: [{ state: "frozen", ySplit: 2 }],
  pageSetup: {
    orientation: "landscape",
    horizontalCentered: true,
    fitToPage: true,
    fitToWidth: 1,
    fitToHeight: 0,
    printTitlesRow: "1:2",
    printArea: "A1:D5",
  },
});

inventory.addRow(["Product inventory"]);
inventory.mergeCells("A1:D1");
inventory.getCell("A1").font = { bold: true, size: 16 };
inventory.addRow(["SKU", "Product", "Updated on", "Stock"]);
inventory.getRow(2).font = { bold: true };
inventory.getRow(2).eachCell((cell) => {
  cell.border = { top: { style: "thin" }, bottom: { style: "double" } };
});
inventory.getRow(2).height = 24;
inventory.getCell("D2").fill = {
  type: "pattern",
  pattern: "solid",
  fgColor: { argb: "FFE2EFDA" },
};
inventory.addRow(["001234567", "Desk lamp", new Date(Date.UTC(2026, 9, 8)), 120]);
inventory.addRow(["009876543", "Notebook", new Date(Date.UTC(2026, 9, 9)), 80]);
inventory.addRow(["", "Total", null, { formula: "SUM(D3:D4)", result: 200 }]);
inventory.columns = [{ width: 16 }, { width: 28 }, { width: 16 }, { width: 18 }];
inventory.getColumn(3).numFmt = "yyyy-mm-dd";
inventory.getColumn(4).numFmt = "#,##0";
inventory.getColumn(4).alignment = { horizontal: "right" };
inventory.autoFilter = { from: "A2", to: "D4" };
inventory.headerFooter.oddFooter = "Page &P of &N";

const reportBytes = await writeWorkbook(report);
```

Dates are serialized from the JavaScript `Date` timestamp using UTC and the
1900 Excel date system. Import recognizes both 1900 and 1904 date systems.
Excel serial dates carry no timezone; use `Date.UTC` for date-only values when
you want the same calendar date regardless of the host timezone. On import,
number formats determine which non-formula numeric cells become `Date` values.

Formula objects contain formula text without a leading `=` and an optional
cached `result`. MinkExcel does not calculate formulas or update cached results
after inputs change. Supply the result for readers that do not recalculate;
exports request full recalculation when opened in Excel by default. Imported
formula results that represent dates remain numeric serials.

## Supported scope and limitations

- Values: strings, numbers, booleans, dates, blanks, errors and cached formulas.
  Strings retain leading zeros, Unicode, whitespace and Excel escape sequences.
- Export: workbook metadata including independent `lastModifiedBy`, multiple
  sheets, row gaps, widths, heights, fonts, solid fills, alignment, colored
  borders, number formats, merged cells, frozen rows, filters, print titles/areas,
  page setup/margins, horizontal print centering and footers.
- Import: shared/inline strings, rich text flattened to text, hyperlink display
  text, cached formulas and 1900/1904 dates. Default imports do not retain
  presentation; template mode retains the original package for value edits.
- Archive handling: ZIP32 STORE and DEFLATE, CRC32 checks, internal relationship
  resolution, XML parsing without external DTD/entity resolution, size and
  nesting limits, and cancellation checkpoints during archive/row processing.
  External relationships are never fetched.

MinkExcel does not read legacy `.xls`, support ZIP64/encryption, implement charts,
images, macros, tables, data validation or conditional formatting, or provide the
complete ExcelJS API. Existing unsupported parts are retained in template mode
but cannot be created or edited through the model. Default imports discard them.

Export uses the standard `CompressionStream("deflate")` API and removes its
zlib envelope to obtain raw ZIP DEFLATE. ZIP headers and checksums are owned by
the package. Import accepts stored and compressed workbooks.

### Import limits and errors

`readWorkbook(bytes, options?, signal?)` accepts `ReadOptions`: the partial
`ReadLimits` object plus an optional `preserveTemplate` boolean (default false).
Fields you omit retain their defaults; byte limits use bytes, not decimal MB.
Explicit limits must be finite nonnegative safe integers, except
`compressionRatio`, which must be finite and greater than zero. An explicit
`undefined` retains the field's default.

| Option | Default | Scope |
| --- | ---: | --- |
| `fileBytes` | 5 MiB (5,242,880 bytes) | Input file |
| `entries` | 1,000 | Archive entries |
| `entryBytes` | 20 MiB (20,971,520 bytes) | Uncompressed data per archive entry |
| `totalBytes` | 50 MiB (52,428,800 bytes) | Total uncompressed archive data |
| `compressionRatio` | 200 | Uncompressed/compressed size per entry |
| `rows` | 10,000 | Maximum worksheet row number minus one |
| `columns` | 100 | Maximum worksheet column number |
| `cells` | 100,000 | Rectangular extent per worksheet |

The row limit allows positions through row 10,001 to accommodate a header.
The reader does not identify a header automatically. Gaps count toward limits:
the cell extent uses the highest parsed row number multiplied by the highest
column number encountered, with at least one column for blank rows.

Limits apply together. A worksheet with a header and 1,000 data rows across
100 columns has an extent of 100,100 cells, exceeding the default cell limit.
Increase `cells` explicitly when that layout is expected:

```ts
import { readWorkbook, XlsxError, XlsxLimitError } from "minkexcel";

try {
  const imported = await readWorkbook(bytes, {
    rows: 1_000,
    columns: 100,
    cells: 100_100,
  });
  console.log(imported.worksheets.length);
} catch (error) {
  if (error instanceof XlsxLimitError) {
    console.error("Workbook exceeds an import limit:", error.message);
  } else if (error instanceof XlsxError) {
    console.error("Workbook is invalid or unsupported:", error.message);
  } else {
    throw error;
  }
}
```

Inflate output is bounded by the entry's declared size and validated against
its CRC32. Import limits do not apply to export.

### Edit an existing template

```ts
const book = await readWorkbook(templateBytes, { preserveTemplate: true });
const sheet = book.getWorksheet(1)!;
book.creator = "PCSTI ERP";
book.modified = new Date();

for (let row = 2; row <= sheet.rowCount; row++) {
  for (let column = 1; column <= 3; column++) sheet.getCell(row, column).value = null;
}
sheet.getCell(2, 1).value = "Employee name";
sheet.getCell(2, 2).value = "000012345678";
sheet.getCell(2, 3).value = 1234.50;
const output = await writeWorkbook(book);
```

Template mode supports clearing, replacing and appending cell values and updating
existing core metadata fields. Original styles, column widths, row heights,
merges, comments, protection, relationships and binary parts are retained. New
cells inherit an existing column style. Presentation remains in the retained
XML; imported style and layout getters do not reconstruct it. Edit the top-left
cell of an existing merge. Dates require an existing date-formatted cell or
column and retain the template's 1900/1904 date system.

Only changed metadata fields are rewritten; missing, unchanged fields do not
prevent export. Changing a metadata field that is absent from the template
throws `XlsxError`.

Adding, removing or reordering worksheets, changing presentation or calculation
properties, editing merged member cells or shared/array formula ranges, and
assigning dates without an existing date format throw `XlsxError`. Formula caches
are not recalculated; changed values request full recalculation when opened in
Excel. The original shared strings and untouched cells remain intact; replacement
strings use inline strings. All retained archive parts are checksum-validated.
This mode retains more data in memory than the default value import. External
relationships may be retained but are never fetched.

### Iteration and row counts

`getWorksheet` accepts an exact name or a one-based numeric worksheet position.
`actualRowCount` counts populated rows; `rowCount` is the furthest created row.
Use `eachRow` for sparse imports instead of treating `actualRowCount` as the last
row number:

```ts
sheet.eachRow((row, number) => {
  if (number > 1) console.log(row.values);
});
sheet.getColumn(3).eachCell((cell, rowNumber) => {
  cell.alignment = { ...cell.alignment, wrapText: true };
});
// Include null cells through rowCount when required:
sheet.getColumn(3).eachCell({ includeEmpty: true }, (cell) => {
  cell.border = { bottom: { style: "thin" } };
});
```

### Cancellation

Both operations accept an optional `AbortSignal`:

```ts
const controller = new AbortController();
const signal = controller.signal;

await readWorkbook(bytes, {}, signal);
await writeWorkbook(book, signal);

// Call controller.abort() from your application's cancel action.
```

An aborted operation rejects with `signal.reason`. Cancellation is checked at
processing checkpoints; use a Web Worker when main-thread responsiveness matters.

## MinkExcel vs ExcelJS

Choose MinkExcel for simple XLSX imports and exports, a compact API and strict
file-integrity checks. Migrating common ExcelJS workflows mainly involves
changing the I/O calls below and checking the supported features. Choose ExcelJS
when you need its broader workbook model, richer formatting or Node streaming I/O.

| ExcelJS | MinkExcel |
| --- | --- |
| `new ExcelJS.Workbook()` | `new Workbook()` |
| `await book.xlsx.writeBuffer()` | `await writeWorkbook(book)` |
| `await book.xlsx.load(bytes)` | `const book = await readWorkbook(bytes)` |
| Load/edit/save an existing template | `readWorkbook(bytes, { preserveTemplate: true })`, then `writeWorkbook(book)` |
| `getWorksheet(1)`, `actualRowCount`, `getColumn(n).eachCell(...)`, `cell.border` | Supported |
| `book.addWorksheet(name)`, `sheet.addRow(values)`, `sheet.getCell(address)` | Same method names for the supported subset |

`readWorkbook` creates a new workbook. Its default import extracts values;
template mode retains original parts for value edits. MinkExcel keeps the
top-level I/O functions and does not expose `workbook.xlsx.load/writeBuffer`.

This comparison uses MinkExcel 0.1.0 and ExcelJS 4.4.0, the version used by the
benchmark. ExcelJS features are documented in its
[versioned README](https://github.com/exceljs/exceljs/blob/v4.4.0/README.md)
and dependencies in its
[manifest](https://github.com/exceljs/exceljs/blob/v4.4.0/package.json).

| Capability | MinkExcel | ExcelJS 4.4.0 |
| --- | --- | --- |
| Runtime npm dependencies | None | Nine direct dependencies, plus transitive dependencies |
| File formats | XLSX | XLSX and CSV |
| Imported presentation | Default: values only. Template mode: retain original parts for cell value edits | Reads supported styles and workbook features for editing; not lossless for arbitrary files |
| Export formatting | Fonts, solid fills, alignment, colored borders, number formats, merges and print settings | Broader styling, including gradient fills |
| Rich text and hyperlinks | Imports display text only | Rich text and hyperlink cell values |
| Images, tables, data validation and conditional formatting | Retained opaquely in templates; no creation/editing API | Supported, with feature-specific limitations |
| Formulas | Formula text and cached results; no calculation engine | Formula text and supplied results; no calculation engine |
| Workbook processing | In memory; optional `AbortSignal` and configurable import limits | In-memory document model and Node streaming reader/writer |
| Browser use | ES modules; export requires `CompressionStream('deflate')` and Web Streams | Browser bundles for the document model; streaming reader/writer excluded |

### Measured performance

MinkExcel was faster for both operations at every tested size (100, 1,000 and
10,000 data rows) across numeric, text and mixed workloads on Bun and Node.js.
At 10,000 rows, exports were about **1.9–2.8× faster** and imports **1.2–1.5×
faster** than ExcelJS 4.4.0 in this run:

| Runtime | Workload, 10,000 rows | MinkExcel export ms | ExcelJS export ms | MinkExcel import ms | ExcelJS import ms |
| --- | --- | ---: | ---: | ---: | ---: |
| Bun 1.4.0 | Numeric | 56.8 | 114.5 | 68.5 | 96.3 |
| Bun 1.4.0 | Text | 70.1 | 130.9 | 85.5 | 121.7 |
| Bun 1.4.0 | Mixed | 69.9 | 133.7 | 79.9 | 118.1 |
| Node 24.12.0 | Numeric | 60.7 | 167.7 | 64.1 | 85.5 |
| Node 24.12.0 | Text | 66.6 | 186.1 | 80.9 | 94.2 |
| Node 24.12.0 | Mixed | 69.3 | 192.2 | 75.5 | 93.4 |

Measured October 8, 2026 on an Apple M4 Pro. Each worksheet has eight columns
plus a header. Times are medians of eleven samples after one warmup, using the
built ES modules without an AbortSignal. Export includes workbook construction
and serialization with the same supported styles. Both readers receive the
same compressed ExcelJS output. Every data cell is checked with both readers
outside timings, including a MinkExcel roundtrip.

MinkExcel imports values while ExcelJS builds a richer presentation model.
Both writers use DEFLATE, with different compression settings and workbook XML.
These document-model measurements do not compare streaming, browser speed or
peak memory, and are not universal performance guarantees. Full results,
raw samples and file sizes: [Bun](benchmarks/MATRIX.md)
([JSON](benchmarks/matrix.json)), [Node](benchmarks/MATRIX-NODE.md)
([JSON](benchmarks/matrix-node.json)). See
[CONTRIBUTING.md](https://github.com/accntech/minkexcel/blob/main/CONTRIBUTING.md)
to reproduce them.

### File integrity and text preservation

The following checks use the same input values or file bytes with both
libraries. String preservation checks each writer's output with both readers;
file checks use the default public import options.

| Check | MinkExcel | ExcelJS 4.4.0 |
| --- | --- | --- |
| Preserve literal `_x0041_` and `_x005F_` strings | Preserved | Changed |
| Preserve a carriage return within cell text | Preserved | Changed |
| Preserve Unicode, XML characters and surrounding spaces | Preserved | Preserved |
| Read a worksheet with a mismatched declared CRC32 | Rejected | Accepted |
| Read a truncated ZIP archive | Rejected | Rejected |
| Read XML that rebinds the reserved `xml` prefix | Rejected | Accepted |
| Read a worksheet containing a DTD | Rejected | Accepted |

MinkExcel's reader validates CRC32, enforces XML namespace rules and rejects
DTDs. Configurable import bounds also reject invalid limit values instead of
silently disabling checks. These results cover specific cases; ExcelJS has
broader workbook support, and reliability still depends on the features and
files your application uses. See the [reproducible checks](benchmarks/RELIABILITY.md)
and [recorded values](benchmarks/reliability.json).

## Development and releases

All participants are expected to follow the
[Code of Conduct](https://github.com/accntech/minkexcel/blob/main/CODE_OF_CONDUCT.md).

See [CONTRIBUTING.md](https://github.com/accntech/minkexcel/blob/main/CONTRIBUTING.md)
for build, test, browser harness and benchmark commands, and
[RELEASING.md](https://github.com/accntech/minkexcel/blob/main/RELEASING.md)
for npm publishing and GitHub release setup.

## License

Released under the [MIT License](LICENSE).
