import { icon } from "./icons.mjs";
export const escape = value => String(value).replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
function highlight(text, language) {
  if (!language.includes('TypeScript') && language !== 'Signatures') return escape(text);
  const tokens = /("(?:\\.|[^"\\])*"|'(?:\\.|[^'\\])*'|`[^`]*`)|(\/\/[^\n]*)|\b(import|from|export|const|let|new|await|async|function|return|if|else|throw|try|catch|for|of|type|interface|extends|readonly|true|false|null|undefined|void|number|string|boolean)\b|\b(\d+(?:\.\d+)?)\b/g;
  let output = '', position = 0;
  for (const match of text.matchAll(tokens)) {
    output += escape(text.slice(position, match.index));
    const kind = match[1] ? 'string' : match[2] ? 'comment' : match[3] ? 'keyword' : 'number';
    output += `<span class="syntax-${kind}">${escape(match[0])}</span>`;
    position = match.index + match[0].length;
  }
  return output + escape(text.slice(position));
}
export const code = (text, language = 'TypeScript') => `<div class="code-block"><div class="code-heading"><span>${language}</span><button type="button" class="copy" aria-label="Copy code" hidden>${icon('copy')}<span>Copy</span></button></div><pre tabindex="0" aria-label="${language} code"><code>${highlight(text.trim(), language)}</code></pre></div>`;
const note = (title, text) => `<aside class="note">${icon("document-text")}<div><strong>${title}</strong><p>${text}</p></div></aside>`;
const table = (headers, rows) => `<div class="table-wrap" tabindex="0" role="region" aria-label="Data table; scroll horizontally if needed"><table role="table"><thead><tr>${headers.map(h => `<th scope="col">${h}</th>`).join('')}</tr></thead><tbody>${rows.map(row => `<tr>${row.map((cell, i) => `<${i ? 'td' : 'th scope="row"'}>${cell}</${i ? 'td' : 'th'}>`).join('')}</tr>`).join('')}</tbody></table></div>`;
const section = (id, title, body) => `<section id="${id}"><h2>${title}<a class="anchor" href="#${id}" aria-label="Link to ${title}">#</a></h2>${body}</section>`;
const signature = text => `<code>${escape(text)}</code>`;
const apiTable = rows => table(['Member', 'Behavior'], rows.map(([member, behavior]) => [signature(member), behavior])).replace('class="table-wrap"', 'class="table-wrap api-table"');
export function pages({ version, bun, node, timingChart, sizeChart, resultsTable, types }) {
  const first = bun.rows.filter(r => r.rows === 10000);
  const ratios = [...bun.rows, ...node.rows].filter(r => r.rows === 10000).map(r => r.exceljsExportMs / r.minkexcelExportMs);
  return [
    { slug: 'index', label: 'Introduction', title: 'Introduction', description: 'MinkExcel is a small, dependency-free library for reading XLSX values and creating formatted spreadsheet reports in JavaScript.', category: 'Overview', body: `
      <div class="hero-actions"><a class="button primary" href="getting-started.html">Getting started ${icon("arrow-right")}</a><a class="text-link" href="api.html">API reference ${icon("arrow-right")}</a></div>
      ${code('npm install minkexcel', 'Terminal')}
      <div class="package-facts"><span>${icon('check-circle')} Zero runtime dependencies</span><span>${icon('code-square')} ESM + TypeScript</span><span>${icon('download')} 11.0 KiB gzip¹</span></div>
      ${section('why', 'Why this project exists', `<p>Many applications need a straightforward spreadsheet workflow: export a data grid, create a formatted report, or read uploaded rows. A broad workbook library can bring more code and a larger feature surface than those jobs require.</p><p>MinkExcel was developed to give those common ExcelJS workflows a smaller, focused implementation. It keeps familiar worksheet methods, owns its ZIP and XML processing, and uses standard web APIs for export compression. The result is a library with no runtime npm dependencies that works in browsers, module Web Workers, Bun and Node.js.</p><p>Its priorities are fast processing, predictable cell values and dependable file handling. Leading-zero identifiers stay strings, literal Excel escape sequences stay literal, and imports validate archive checksums with explicit resource bounds.</p>`)}
      ${section('principles', 'A small surface, by design', `<div class="principles"><article>${icon("upload", "mini-icon")}<h3>Export application data</h3><p>Build new XLSX reports with fonts, number formats, merges, frozen headings and print settings.</p></article><article>${icon("download", "mini-icon")}<h3>Read the values you need</h3><p>Extract strings, numbers, booleans, dates, errors and cached formulas from uploaded workbooks.</p></article><article>${icon("shield-check", "mini-icon")}<h3>Check the file boundary</h3><p>CRC32 validation, strict XML handling, bounded decompression and configurable import limits.</p></article></div>`)}
      ${note('Choose the right scope', 'Import is value-oriented. Reading and rewriting a workbook does not preserve its styles, layout or unsupported features. Use ExcelJS when your application needs a broader workbook model or Node streaming I/O. See the <a href="comparison.html">implementation comparison</a>.')}
      ${section('next', 'From data to a workbook', code(`import { Workbook, writeWorkbook } from "minkexcel";

const book = new Workbook();
const sheet = book.addWorksheet("Products");
sheet.addRow(["SKU", "Product", "Price"]);
sheet.addRow(["001234567", "Desk lamp", 1234.56]);
sheet.getColumn(3).numFmt = "#,##0.00";

const bytes = await writeWorkbook(book); // Uint8Array`)+`<p>Start with the <a href="getting-started.html">export and import guide</a>, or inspect the <a href="benchmarks.html">recorded benchmark results</a>.</p><p class="fine-print">¹ Full browser ESM bundle: 30,624 minified bytes / 11,301 gzip bytes. Measured with Bun 1.4.0 on October 8, 2026. Size varies with bundler and imports.</p>`)}
    ` },
    { slug: 'getting-started', label: 'Getting started', title: 'Getting started', description: 'Create a report, turn it into XLSX bytes, and read the values back.', category: 'Guides', body: `
      ${section('create', 'Create and export', `<p>Install the package with <code>npm install minkexcel</code>. Use named ES module imports, create a workbook and add a worksheet. Your application decides where the resulting bytes go.</p>` + code(`import { Workbook, writeWorkbook } from "minkexcel";

const book = new Workbook();
const products = book.addWorksheet("Products");
products.addRow(["SKU", "Product", "Price"]);
products.addRow(["001234567", "Desk lamp", 1234.56]);
products.addRow(["009876543", "Notebook", 80]);
products.getRow(1).font = { bold: true };
products.columns = [{ width: 16 }, { width: 28 }, { width: 18 }];
products.getColumn(3).numFmt = "#,##0.00";

const bytes = await writeWorkbook(book);`))}
      ${section('read', 'Import and read values', code(`import { readWorkbook } from "minkexcel";

const imported = await readWorkbook(bytes);
const products = imported.getWorksheet("Products");
if (!products) throw new Error("Products worksheet is missing.");

console.log(products.getCell("A2").value); // "001234567"
console.log(products.getCell(2, 3).value);  // 1234.56

products.eachRow((row, number) => {
  if (number === 1) return; // Skip this report's header.
  console.log(row.getCell(1).value, row.getCell(3).value);
});`) + `<p><code>readWorkbook</code> creates a new workbook. Import extracts values and cached formula results; presentation is not restored.</p>`)}
      ${section('browser', 'Browser files and downloads', code(`import { readWorkbook, writeWorkbook, type Workbook } from "minkexcel";

async function importFile(file: File): Promise<Workbook> {
  return readWorkbook(new Uint8Array(await file.arrayBuffer()));
}

async function downloadWorkbook(book: Workbook) {
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
}`))}
      ${section('node', 'Read and save in Node.js', code(`import { readFile, writeFile } from "node:fs/promises";
import { Workbook, readWorkbook, writeWorkbook } from "minkexcel";

const book = new Workbook();
book.addWorksheet("Products").addRow(["001234567", 1234.56]);
await writeFile("products.xlsx", await writeWorkbook(book));

const imported = await readWorkbook(new Uint8Array(await readFile("products.xlsx")));
console.log(imported.getWorksheet("Products")?.getCell("A1").value);`))}
      ${section('conventions', 'Know the indexing rules', `<p>Row and column positions are one-based: <code>getCell(2, 3)</code> is <code>C2</code>. Cell addresses use uppercase letters. Arrays passed to <code>addRow</code> or assigned to <code>row.values</code> are zero-based, while reading <code>row.values</code> returns a one-based sparse array.</p>` + code(`const row = products.addRow(["001234567", "Desk lamp", 1234.56]);
console.log(row.values[0]); // undefined
console.log(row.values[1]); // "001234567"
console.log(row.values[3]); // 1234.56
const values = row.values.slice(1);`) + `<p><code>getRow</code> and <code>getCell</code> create missing entries. Iteration skips blank rows and cells. Counts describe the furthest created positions. Keep identifiers as strings to preserve leading zeros.</p>`)}
      ${section('report', 'Add report formatting', code(`const report = new Workbook();
report.creator = "Product team";
const sheet = report.addWorksheet("Inventory", {
  views: [{ state: "frozen", ySplit: 2 }],
  pageSetup: {
    orientation: "landscape", fitToPage: true,
    fitToWidth: 1, fitToHeight: 0,
    printTitlesRow: "1:2", printArea: "A1:D5",
  },
});
sheet.addRow(["Product inventory"]);
sheet.mergeCells("A1:D1");
sheet.getCell("A1").font = { bold: true, size: 16 };
sheet.addRow(["SKU", "Product", "Updated", "Stock"]);
sheet.getRow(2).font = { bold: true };
sheet.addRow(["001234567", "Desk lamp", new Date(Date.UTC(2026, 9, 8)), 120]);
sheet.addRow(["009876543", "Notebook", new Date(Date.UTC(2026, 9, 9)), 80]);
sheet.addRow(["", "Total", null, { formula: "SUM(D3:D4)", result: 200 }]);
sheet.getColumn(3).numFmt = "yyyy-mm-dd";
sheet.getColumn(4).numFmt = "#,##0";
sheet.getCell("D2").fill = {
  type: "pattern", pattern: "solid", fgColor: { argb: "FFE2EFDA" },
};
sheet.autoFilter = { from: "A2", to: "D4" };
sheet.headerFooter.oddFooter = "Page &P of &N";
const reportBytes = await writeWorkbook(report);`) + note('Dates and formulas', 'Export serializes Date timestamps using UTC and the 1900 date system; import recognizes 1900 and 1904. Use Date.UTC for date-only values. Formula text has no leading =. Supply cached results yourself: MinkExcel does not calculate formulas. Imported formula date results stay numeric serials.'))}
      ${section('workers', 'Move larger jobs into a Web Worker', `<p>Workbooks are processed in memory. For a responsive browser interface, use a module Web Worker and transfer the XLSX byte buffer back to the main thread. Bundle the worker through your application tooling.</p>` + code(`// workbook.worker.ts
import { Workbook, writeWorkbook } from "minkexcel";

self.onmessage = async (event) => {
  try {
    const book = new Workbook();
    const sheet = book.addWorksheet("Data");
    for (const values of event.data) sheet.addRow(values);
    const bytes = await writeWorkbook(book);
    self.postMessage({ bytes }, { transfer: [bytes.buffer] });
  } catch (error) {
    self.postMessage({ error: String(error) });
  }
};`) + code(`const worker = new Worker(new URL("./workbook.worker.ts", import.meta.url), {
  type: "module",
});
worker.onmessage = ({ data }) => {
  if (data.error) console.error(data.error);
  else console.log(data.bytes); // Save or download the bytes.
  worker.terminate();
};
worker.postMessage([["SKU", "Price"], ["001234567", 1234.56]]);`))}
    ` },
    { slug: 'installation', label: 'Installation', title: 'Installation', description: 'One package. Named imports. No runtime dependencies.', category: 'Guides', body: `
      ${section('install', 'Install the package', code('npm install minkexcel', 'npm') + code('bun add minkexcel', 'Bun') + code('pnpm add minkexcel', 'pnpm') + `<p>The package ships ES modules and TypeScript declarations. It is ESM-only; use <code>import</code>. No separate type package is needed.</p>` + code('import { Workbook, readWorkbook, writeWorkbook } from "minkexcel";'))}
      ${section('runtime', 'Runtime requirements', `<p>Export requires <code>CompressionStream("deflate")</code> and Web Streams. Import uses the package’s own inflater and does not require <code>DecompressionStream</code>. The library performs no filesystem or network I/O.</p>` + table(['Environment', 'Verified version', 'Verification'], [['Bun', '1.4.0', 'Examples, unit tests, interoperability and benchmarks'], ['Node.js', '24.12.0', 'Examples and benchmarks using built ESM'], ['Chromium', '153.0.8010.12', 'Browser bundle and module Web Worker import/export']]) + `<p class="fine-print">Verified October 8, 2026. These are tested versions, not minimum supported versions. Validate other runtimes and browsers in your application.</p>`)}
      ${section('tree-shaking', 'Import only what you use', code('import { Workbook, writeWorkbook } from "minkexcel";') + `<p>ES modules and <code>sideEffects: false</code> let application bundlers remove unused exports. Export-only builds can exclude the reader and inflater. Model-only builds can exclude ZIP and XML processing.</p>` + table(['Public imports', 'Minified bytes', 'Gzip bytes'], [['Workbook', '4,532', '1,763'], ['Workbook + writeWorkbook', '16,630', '6,090'], ['readWorkbook', '19,150', '7,513'], ['All public exports', '30,624', '11,301']]) + `<p>Measured with Bun 1.4.0 on October 8, 2026, excluding declarations and the icon. Methods on retained model classes generally remain. Direct browser or Node imports without bundling do not remove unused files.</p>`)}
      ${section('source', 'Build from source', code(`git clone https://github.com/accntech/minkexcel.git
cd minkexcel
bun install --cwd tools --frozen-lockfile --ignore-scripts
bun run build
bun run test`, 'Terminal') + `<p>Optional contributor tools are isolated in <code>tools/</code>. The build writes ES modules and declarations to <code>dist/</code>. Read the <a href="https://github.com/accntech/minkexcel/blob/main/CONTRIBUTING.md">contributor guide</a> for browser, interoperability and package checks.</p>`)}
      ${section('scope', 'Supported file scope', `<p>MinkExcel handles XLSX ZIP32 archives with STORE or DEFLATE compression. It does not read legacy XLS, ZIP64 or encrypted workbooks. Charts, images, macros, tables, data validation and conditional formatting are outside its supported scope.</p>` + note('Import boundaries', 'The default file limit is 5 MiB. Row, column, cell and decompression limits also apply together. See <a href="api.html#read-limits">ReadLimits</a> before accepting larger uploads.'))}
    ` },
    { slug: 'api', label: 'API reference', title: 'API reference', description: 'The complete public surface for MinkExcel '+version+'. All row and column positions are one-based.', category: 'Reference', body: `
      ${section('io', 'Read and write', code(`writeWorkbook(book: Workbook, signal?: AbortSignal): Promise<Uint8Array>
readWorkbook(bytes: Uint8Array, limits?: ReadLimits, signal?: AbortSignal): Promise<Workbook>`, 'Signatures') + `<p><code>writeWorkbook</code> serializes a new XLSX archive in memory. <code>readWorkbook</code> returns a new workbook of imported values and cached formulas. Your application supplies and saves the bytes. Import limits do not apply to export.</p>`)}
      ${section('workbook', 'Workbook', apiTable([
        ['new Workbook()', 'Create an empty workbook.'], ['worksheets: Worksheet[]', 'Readonly array reference containing sheets in insertion order.'], ['creator: string', 'Author metadata; defaults to MinkExcel.'], ['created: Date; modified: Date', 'Created and modified timestamps; both initially reference the creation date.'], ['calcProperties: { fullCalcOnLoad: boolean }', 'Request Excel recalculation on open; defaults to true.'], ['addWorksheet(name, options?): Worksheet', 'Create a sheet. Name must be 1–31 characters, exclude \\ / * ? : [ ], not start/end with an apostrophe and be unique ignoring case.'], ['getWorksheet(name): Worksheet | undefined', 'Exact, case-sensitive lookup by name.']]))}
      ${section('worksheet', 'Worksheet', apiTable([
        ['new Worksheet(name, options?)', 'Direct constructor; prefer Workbook.addWorksheet for name validation and registration.'], ['name: string', 'Readonly worksheet name.'], ['addRow(values: unknown[]): Row', 'Append after the furthest created row; consumes a zero-based array of supported values.'], ['getRow(number): Row', 'Get or create a row. Range: 1–1,048,576.'], ['getCell(address: string): Cell', 'Get or create a cell using an uppercase A1 address.'], ['getCell(row: number, column: number): Cell', 'Get or create a cell by position.'], ['getColumn(number): Column', 'Get or create column metadata. Range: 1–16,384.'], ['columns = [{ width: number }, …]', 'Set widths from a zero-based array. Setter only; no column keys or headers API.'], ['rowCount; columnCount: number', 'Furthest created row and cell column, not populated-entry counts. Column definitions alone do not increase columnCount.'], ['eachRow((row, number) => void)', 'Visit populated rows in ascending position order.'], ['mergeCells(range: string): void', 'Merge an A1 range, for example A1:D1. Non-master values resolve through the top-left cell.'], ['mergeCells(top, left, bottom, right): void', 'Merge by positions. Invalid or overlapping ranges throw XlsxError.'], ['views?: WorksheetOptions["views"]', 'Frozen row views; only frozen state and ySplit are supported.'], ['pageSetup: PageSetup', 'Orientation, sizing, print titles/area and margins.'], ['headerFooter: { oddFooter?: string }', 'Set footer text; Excel placeholders include &amp;P and &amp;N.'], ['autoFilter?: { from; to }', 'Filter endpoints as uppercase addresses or { row, column } objects.'], ['rows; columnDefinitions; merges', 'Readonly collection references: Map&lt;number, Row&gt;, Map&lt;number, Column&gt;, and string[] respectively. Use model methods to maintain positions.']]))}
      ${section('row', 'Row', apiTable([
        ['new Row(sheet: Worksheet, number: number)', 'Create a row object; use worksheet methods to register it.'], ['sheet; number', 'Readonly owning worksheet and one-based row position.'], ['getCell(column: number): Cell', 'Get or create a cell.'], ['values = unknown[]', 'Replace all cells using a zero-based array. Null/undefined entries are skipped. Unsupported values throw XlsxError.'], ['values: CellValue[]', 'Read a one-based sparse array; position zero is unused.'], ['cellCount: number', 'Furthest created cell column.'], ['eachCell((cell, column) => void)', 'Visit non-null cells in ascending column order.'], ['height?: number', 'Export row height in points.'], ['style; font; alignment; numFmt', 'Set row styling, inherited by its cells.'], ['cells: Map<number, Cell>', 'Readonly reference to the cell collection.']]))}
      ${section('cell', 'Cell', apiTable([
        ['new Cell(row: Row, column: number)', 'Create a cell object; use getCell to validate and register it.'], ['row; column; address', 'Readonly parent row, column position, and computed uppercase address.'], ['value: CellValue', 'Read/write content. New cells start as null. Merged cells read/write the master value.'], ['text: string', 'Plain string representation or cached formula result; does not apply numFmt.'], ['style: Style', 'Cell’s own style.'], ['resolvedStyle: Style', 'Shallow precedence: column → row → merge master → cell.'], ['font; alignment; numFmt', 'Convenient style accessors. Reading font/alignment creates a local copy of resolved settings.'], ['fill = Fill', 'Write-only convenience setter for a solid pattern fill. Read through style.fill or resolvedStyle.fill.'], ['master?: Cell', 'Merge-master reference for non-master merged cells.']]))}
      ${section('column', 'Column', apiTable([
        ['new Column(number: number)', 'Create column metadata; prefer sheet.getColumn to register it.'], ['number; letter', 'Readonly position and computed letters (1 → A).'], ['width?: number', 'Excel column width units.'], ['style; font; alignment; numFmt', 'Column styles inherited by cells. Use style.fill for column fills.']]))}
      ${section('types', 'Values and formatting types', `<p>These types are exported by the package. Colors use eight-digit ARGB strings, for example <code>FF217346</code>. Solid fills are the only supported fill kind.</p>` + code(types, 'TypeScript · public types') + `<p>Dates are serialized using UTC. Formula result values are cached, not calculated. Fonts, fills, alignment and number formats apply on export; importing does not restore styles.</p>`)}
      ${section('read-limits', 'ReadLimits', `<p>A partial object: omitted or explicitly undefined fields retain defaults. Byte/count limits must be finite nonnegative safe integers; <code>compressionRatio</code> must be finite and greater than zero.</p>` + table(['Field', 'Default', 'Bound'], [['fileBytes', '5,242,880 (5 MiB)', 'Input file bytes'], ['entries', '1,000', 'ZIP entries'], ['entryBytes', '20,971,520 (20 MiB)', 'Uncompressed bytes per entry'], ['totalBytes', '52,428,800 (50 MiB)', 'Total uncompressed bytes'], ['compressionRatio', '200', 'Uncompressed/compressed size per entry'], ['rows', '10,000', 'Maximum worksheet row number minus one'], ['columns', '100', 'Maximum worksheet column number'], ['cells', '100,000', 'Rectangular extent per sheet']]) + `<p>The row limit permits positions through 10,001 to accommodate a header; no header is detected automatically. Gaps count toward the rectangular extent: highest row × highest column, with at least one column for blank rows. Limits apply together.</p>` + code(`const imported = await readWorkbook(bytes, {
  rows: 1_000,
  columns: 100,
  cells: 100_100, // Header + 1,000 rows × 100 columns.
});`))}
      ${section('errors', 'Errors and cancellation', `<p><code>XlsxError extends Error</code> represents invalid or unsupported input/output. <code>XlsxLimitError extends XlsxError</code> represents an import bound being exceeded. Both have an inherited readonly <code>_tag = "XlsxError"</code>. Check the specific subclass first.</p>` + code(`import { readWorkbook, XlsxError, XlsxLimitError } from "minkexcel";

try {
  const book = await readWorkbook(bytes);
  console.log(book.worksheets.length);
} catch (error) {
  if (error instanceof XlsxLimitError) {
    console.error("Import limit exceeded:", error.message);
  } else if (error instanceof XlsxError) {
    console.error("Invalid or unsupported workbook:", error.message);
  } else {
    throw error;
  }
}`) + `<p>Both operations accept an optional <code>AbortSignal</code>. Cancellation is checked at processing checkpoints and rejects with <code>signal.reason</code>. Use a worker when main-thread responsiveness matters.</p>` + code(`const controller = new AbortController();
const pending = readWorkbook(bytes, {}, controller.signal);
// From your application's cancel action:
controller.abort();
try { await pending; } catch (error) { console.log(error); }

// Export also accepts a signal:
await writeWorkbook(book, new AbortController().signal);`))}
    ` },
    { slug: 'comparison', label: 'ExcelJS comparison', title: 'ExcelJS comparison', description: 'Compare the supported API and migrate common workflows from ExcelJS 4.4.0.', category: 'Reference', body: `
      ${section('migration', 'Change the I/O boundary', `<p>MinkExcel keeps common worksheet method names. Replace workbook creation and XLSX I/O, then check every feature your application uses. This is a supported subset of ExcelJS, not a complete replacement.</p>` + table(['ExcelJS 4.4.0', 'MinkExcel '+version], [['<code>new ExcelJS.Workbook()</code>', '<code>new Workbook()</code>'], ['<code>await book.xlsx.writeBuffer()</code>', '<code>await writeWorkbook(book)</code>'], ['<code>await book.xlsx.load(bytes)</code>', '<code>const book = await readWorkbook(bytes)</code>'], ['<code>book.addWorksheet(name)</code>', 'Same supported method name'], ['<code>sheet.addRow(values)</code>', 'Same method name; zero-based input array'], ['<code>sheet.getCell("A1")</code>', 'Same method name and address convention']]) + `<div class="code-pair"><div><h3>ExcelJS</h3>${code(`import ExcelJS from "exceljs";
const book = new ExcelJS.Workbook();
book.addWorksheet("Data").addRow(["001234", 42]);
const bytes = await book.xlsx.writeBuffer();`)}</div><div><h3>MinkExcel</h3>${code(`import { Workbook, writeWorkbook } from "minkexcel";
const book = new Workbook();
book.addWorksheet("Data").addRow(["001234", 42]);
const bytes = await writeWorkbook(book);`)}</div></div>`)}
      ${section('differences', 'Implementation differences that matter', `<p>Use numeric column positions in MinkExcel: column keys, header definitions and object-based row insertion are not implemented. Its <code>row.values</code> setter accepts zero-based inputs even though the getter returns a one-based sparse array. Do not assign a copied ExcelJS sparse row array without adapting the indexing.</p><p><code>readWorkbook</code> returns a new workbook and imports values. It does not populate a pre-existing object or retain imported presentation. The library returns <code>Uint8Array</code> bytes and leaves files, network requests and downloads to your application.</p>` + table(['Capability', 'MinkExcel', 'ExcelJS 4.4.0'], [['Runtime npm dependencies', 'None', 'Nine direct dependencies plus transitive dependencies'], ['File formats', 'XLSX', 'XLSX and CSV'], ['Imported presentation', 'Values; no retained styles/layout', 'Supported styles and workbook features; not lossless for arbitrary files'], ['Export styling', 'Fonts, solid fills, alignment, number formats, merges, print settings', 'Broader styling, including borders and gradient fills'], ['Rich text / hyperlinks', 'Import display text only', 'Rich text and hyperlink values'], ['Images / tables / validation', 'Unsupported', 'Supported, with feature limitations'], ['Formulas', 'Text + supplied cached result; no calculation', 'Text + supplied result; no calculation'], ['Processing', 'In memory; import limits and optional AbortSignal', 'Document model and Node streaming reader/writer'], ['Browser', 'ESM; export needs CompressionStream and Web Streams', 'Document-model browser bundles; streaming excluded']]) + `<p class="fine-print">Comparison is pinned to the benchmark’s ExcelJS version. Sources: <a href="https://github.com/exceljs/exceljs/blob/v4.4.0/README.md">ExcelJS 4.4.0 documentation</a> and <a href="https://github.com/exceljs/exceljs/blob/v4.4.0/package.json">package manifest</a>.</p>`)}
      ${section('integrity', 'Recorded integrity checks', `<p>These checks use identical values or input file bytes. Text preservation checks each writer with both readers. File checks use default public import options.</p>` + table(['Check', 'MinkExcel', 'ExcelJS 4.4.0'], [['Literal _x0041_ / _x005F_ text', 'Preserved', 'Changed'], ['Carriage return in cell text', 'Preserved', 'Changed'], ['Unicode, XML characters, spaces', 'Preserved', 'Preserved'], ['Mismatched declared CRC32', 'Rejected', 'Accepted'], ['Truncated ZIP archive', 'Rejected', 'Rejected'], ['Rebound reserved xml prefix', 'Rejected', 'Accepted'], ['Worksheet DTD', 'Rejected', 'Accepted']]) + `<p>MinkExcel validates CRC32, enforces XML namespace rules and rejects DTDs. External relationships are never fetched. These are specific tested cases; reliability depends on the features and files your application uses. <a href="https://github.com/accntech/minkexcel/blob/main/benchmarks/RELIABILITY.md">Read the reproducible checks</a>.</p>`)}
      ${section('performance', 'Compare the measured workloads', `<p>Across the recorded 10,000-row tests, exports were ${Math.min(...ratios).toFixed(1)}–${Math.max(...ratios).toFixed(1)}× faster. The reader scopes differ: MinkExcel reads values while ExcelJS builds a richer presentation model. Explore runtime, workload and size differences on the <a href="benchmarks.html">benchmark page</a>.</p>`)}
    ` },
    { slug: 'benchmarks', label: 'Benchmarks', title: 'Benchmarks', description: 'Recorded measurements, raw samples and the context to interpret them.', category: 'Measurements', body: `
      <div class="benchmark-summary"><span class="summary-label">Recorded export comparison · 10,000 rows</span><strong>${Math.min(...ratios).toFixed(1)}–${Math.max(...ratios).toFixed(1)}×</strong><p>faster export in the recorded workloads</p><span class="fine-print">MinkExcel ${version} vs ExcelJS 4.4.0 · Apple M4 Pro · October 8, 2026</span></div>
      ${section('timings', 'Time to process a workbook', `<p>Switch the runtime, operation or row count to compare the same three workloads. Every bar begins at zero. Lower is better.</p><div class="chart-controls" hidden><label>Runtime<span class="select-field"><select id="runtime"><option value="bun">Bun 1.4.0</option><option value="node">Node 24.12.0</option></select>${icon('alt-arrow-down')}<template>${icon('check')}</template></span></label><label>Operation<span class="select-field"><select id="operation"><option value="Export">Export</option><option value="Import">Import</option></select>${icon('alt-arrow-down')}<template>${icon('check')}</template></span></label><label>Data rows<span class="select-field"><select id="row-count"><option value="100">100</option><option value="1000">1,000</option><option value="10000" selected>10,000</option></select>${icon('alt-arrow-down')}<template>${icon('check')}</template></span></label></div><div class="chart-panel"><div class="chart-top"><h3 id="chart-title">Export · 10,000 data rows</h3><div class="legend"><span class="mink-key">MinkExcel</span><span class="excel-key">ExcelJS</span></div></div><div id="timing-chart">${timingChart(first)}</div><p id="chart-caption" class="chart-caption" aria-live="polite">Bun 1.4.0 · medians of 11 samples after 1 warmup · 8 columns plus a header</p></div><noscript><p>Showing Bun export for 10,000 rows. Download both runtime datasets below for all results.</p></noscript>`)}
      ${section('file-sizes', 'The output footprint', `<p>Compressed XLSX sizes for the selected runtime and row count. Both writers use DEFLATE, with different compression settings and workbook XML. File size is not a measure of peak memory.</p><div class="chart-panel"><div class="chart-top"><h3>Compressed XLSX output</h3><div class="legend"><span class="mink-key">MinkExcel</span><span class="excel-key">ExcelJS</span></div></div><div id="size-chart">${sizeChart(first)}</div></div>`)}
      ${section('results', 'Every recorded result', `<p id="table-caption">Bun 1.4.0 · all workloads and row counts. Times in milliseconds.</p><div class="table-wrap" tabindex="0" role="region" aria-label="Benchmark results; scroll horizontally if needed"><table id="result-table">${resultsTable(bun.rows)}</table></div><div class="download-links"><a href="data/matrix.json" download>${icon("download")} Bun JSON</a><a href="data/matrix-node.json" download>${icon("download")} Node JSON</a></div>`)}
      ${section('method', 'How these numbers were measured', `<dl class="method"><div><dt>Machine</dt><dd>Apple M4 Pro · macOS · arm64</dd></div><div><dt>Runtimes</dt><dd>Bun 1.4.0 and Node.js 24.12.0</dd></div><div><dt>Samples</dt><dd>Median of 11 samples after 1 warmup</dd></div><div><dt>Shape</dt><dd>100, 1,000 and 10,000 data rows · 8 columns + header</dd></div><div><dt>Versions</dt><dd>MinkExcel ${version} · ExcelJS 4.4.0</dd></div><div><dt>Cancellation</dt><dd>No AbortSignal supplied</dd></div></dl><p>Export includes workbook construction and serialization with the same supported styles. Both readers receive the same compressed ExcelJS output. Every data cell is checked with both readers outside timed samples, including a MinkExcel roundtrip. The built package ES modules are used.</p><p><strong>Numeric:</strong> fractions, negatives and zeros. <strong>Text:</strong> leading-zero identifiers, Unicode, XML characters, formula-like text and whitespace. <strong>Mixed:</strong> strings, dates, booleans, numbers, cached formulas, blanks and errors. Literal Excel escape sequences are checked separately by the reliability suite.</p>` + note('Interpret the scope', 'MinkExcel imports values; ExcelJS builds a richer presentation model. These local document-model tests do not compare streaming, browser speed or peak memory and do not establish a universal speed advantage. Raw JSON includes timing samples, exact values and file sizes.'))}
      ${section('reproduce', 'Reproduce the comparison', code(`bun install --cwd tools --frozen-lockfile --ignore-scripts
bun run build
XLSX_BENCH_ITERATIONS=11 bun run bench:matrix
XLSX_BENCH_ITERATIONS=11 npm run bench:node`, 'Terminal') + `<p>Run sequentially on an otherwise idle machine. The Node command needs native TypeScript type stripping. Measurements write to <code>benchmarks/matrix.json</code> and <code>benchmarks/matrix-node.json</code>. Rebuild this documentation to pick up updated data.</p><p>Inspect the <a href="https://github.com/accntech/minkexcel/blob/main/benchmarks/matrix.ts">benchmark implementation</a> and the <a href="https://github.com/accntech/minkexcel/blob/main/CONTRIBUTING.md#benchmarks">reproduction guide</a>.</p>`)}
    ` },
  ];
}
