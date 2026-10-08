# MinkExcel

<img src="assets/icon.png" alt="MinkExcel white mink and XL icon on a solid Excel green squircle" width="160" height="160">

Small XLSX import and export for accounting workflows. This
standalone TypeScript package has **zero npm dependencies**, including development
and peer dependencies. Its source uses standard JavaScript APIs and has no
Node/Bun imports, filesystem access, network access, or application framework imports.
It ships ES modules and TypeScript declarations for browser bundles, Web Workers,
Bun, and modern Node runtimes with the standard compression APIs.

## Usage

```sh
npm install minkexcel
```

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

The package build, browser tests and the ExcelJS benchmark use development tools isolated
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

## Publishing releases

`.github/workflows/publish.yml` publishes to npm when a GitHub release is
published. It checks out the release tag, verifies that the tag matches the
version in `package.json`, runs the unit tests, and builds JavaScript and
TypeScript declarations before publishing. The npm package includes `dist/`,
the README, the manifest and the icon. The library remains dependency-free;
TypeScript is installed only in the separate `tools/` project.

### One-time npm setup

The package must exist on npm before configuring
[trusted publishing](https://docs.npmjs.com/trusted-publishers/). Publish the
initial version from a local checkout with an npm account that owns the package:

```sh
bun install --cwd tools --frozen-lockfile --ignore-scripts
bun run test
npm login
npm publish
```

`npm publish` builds the package through the `prepack` script. Then open the
package's settings on npmjs.com and add a GitHub Actions trusted publisher:

- Organization or user: `accntech`
- Repository: `minkexcel`
- Workflow filename: `publish.yml`
- Environment name: leave blank
- Allowed actions: enable direct publishing with `npm publish`

The workflow authenticates with OIDC, so no npm token or GitHub secret is needed.
Complete the first workflow publish within two days of adding the trusted
publisher; npm expires new configurations that have not yet published.

### Subsequent releases

Update `package.json` to a new, unpublished version and commit it with the
release changes. Publish a GitHub release for that commit using a matching tag,
such as `v0.1.1` for version `0.1.1` (tags without the `v` prefix also work).
The tag must include this workflow and the packaging configuration.

Stable releases publish to npm's `latest` tag. GitHub prereleases and versions
such as `0.2.0-beta.1` publish to `next`. Draft releases and branch pushes do
not publish. For a local package preview, run `npm pack --dry-run` after
installing the tools.
