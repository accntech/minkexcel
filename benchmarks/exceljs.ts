import ExcelJS from "../tools/node_modules/exceljs/excel.js";
import { Workbook, readWorkbook, writeWorkbook } from "../src/index.js";
import { strict as assert } from "node:assert";
import { cpus } from "node:os";

// ExcelJS is a development-only reference installed in the optional tools project.
const counts = [100, 1_000, 10_000];
const iterations = Number(process.env.XLSX_BENCH_ITERATIONS ?? 5);
if (!Number.isInteger(iterations) || iterations < 1)
  throw new Error("XLSX_BENCH_ITERATIONS must be a positive integer.");
const created = new Date("2026-10-08T00:00:00Z");
const headers = [
  "TIN",
  "Customer",
  "Account",
  "Date",
  "Description",
  "Debit",
  "Credit",
  "Balance",
];
function populate(book: Workbook | ExcelJS.Workbook, count: number) {
  book.creator = "MinkExcel";
  book.created = created;
  book.modified = created;
  const sheet = book.addWorksheet("General ledger", {
    views: [{ state: "frozen", ySplit: 1 }],
  });
  sheet.columns = headers.map((_, i) => ({
    width: i === 1 || i === 4 ? 28 : 18,
  }));
  const header = sheet.addRow(headers);
  header.font = { bold: true, size: 11 };
  header.height = 24;
  for (let i = 1; i <= count; i++) {
    sheet.addRow([
      String(i).padStart(9, "0"),
      `Customer ${i % 200}`,
      "1000",
      created,
      `Invoice ${i}`,
      i / 100,
      i % 3 ? 0 : i / 200,
      {
        formula: `F${i + 1}-G${i + 1}`,
        result: i / 100 - (i % 3 ? 0 : i / 200),
      },
    ]);
  }
  for (const i of [6, 7, 8]) sheet.getColumn(i).numFmt = "#,##0.00";
  sheet.getColumn(4).numFmt = "yyyy-mm-dd";
  sheet.pageSetup = {
    orientation: "landscape",
    paperSize: 9,
    fitToPage: true,
    fitToWidth: 1,
    fitToHeight: 0,
    printTitlesRow: "1:1",
  };
  return book;
}
async function median<T>(
  run: () => Promise<T>,
): Promise<{ ms: number; result: T }> {
  await run(); // Warmup excludes module initialization and first-time JIT work.
  const samples: number[] = [];
  let result!: T;
  for (let i = 0; i < iterations; i++) {
    const start = performance.now();
    result = await run();
    samples.push(performance.now() - start);
  }
  samples.sort((a, b) => a - b);
  return { ms: samples[Math.floor(samples.length / 2)], result };
}
const rows = [];
for (const count of counts) {
  const ownExport = await median(() =>
    writeWorkbook(populate(new Workbook(), count) as Workbook),
  );
  const foreignExport = await median(
    async () =>
      new Uint8Array(
        await (
          populate(new ExcelJS.Workbook(), count) as ExcelJS.Workbook
        ).xlsx.writeBuffer(),
      ),
  );
  // Both readers receive the exact same ExcelJS-produced compressed bytes.
  const ownImport = await median(() => readWorkbook(foreignExport.result));
  const foreignImport = await median(async () => {
    const book = new ExcelJS.Workbook();
    await book.xlsx.load(foreignExport.result as unknown as ExcelJS.Buffer);
    return book;
  });
  // Verify exported bytes with the independent reader outside timed samples.
  const verification = new ExcelJS.Workbook();
  await verification.xlsx.load(ownExport.result as unknown as ExcelJS.Buffer);
  for (const book of [ownImport.result, foreignImport.result, verification]) {
    const sheet = book.worksheets[0];
    assert.equal(sheet.rowCount, count + 1);
    assert.equal(sheet.getCell("A2").value, "000000001");
    assert.equal(
      sheet.getCell(`B${count + 1}`).text,
      `Customer ${count % 200}`,
    );
    assert.deepEqual(sheet.getCell(`H${count + 1}`).value, {
      formula: `F${count + 1}-G${count + 1}`,
      result: count / 100 - (count % 3 ? 0 : count / 200),
    });
  }
  assert.equal(
    verification.worksheets[0].getCell("D2").value?.valueOf(),
    created.valueOf(),
  );
  rows.push({
    rows: count,
    minkexcelExportMs: ownExport.ms,
    exceljsExportMs: foreignExport.ms,
    minkexcelImportMs: ownImport.ms,
    exceljsImportMs: foreignImport.ms,
    minkexcelBytes: ownExport.result.length,
    exceljsBytes: foreignExport.result.length,
  });
}
const excelPackage = await Bun.file(
  new URL("../tools/node_modules/exceljs/package.json", import.meta.url),
).json();
const metadata = {
  date: new Date().toISOString(),
  runtime: `Bun ${Bun.version}`,
  platform: `${process.platform} ${process.arch}`,
  cpu: cpus()[0]?.model,
  exceljs: excelPackage.version,
  iterations,
  warmups: 1,
  minkexcelCompression: "DEFLATE via CompressionStream",
};
const lines = rows.map(
  (r) =>
    `| ${r.rows.toLocaleString("en-US")} | ${r.minkexcelExportMs.toFixed(1)} | ${r.exceljsExportMs.toFixed(1)} | ${r.minkexcelImportMs.toFixed(1)} | ${r.exceljsImportMs.toFixed(1)} | ${(r.minkexcelBytes / 1024).toFixed(1)} | ${(r.exceljsBytes / 1024).toFixed(1)} |`,
);
const report = `# XLSX benchmark against ExcelJS\n\n${metadata.date}; ${metadata.runtime}; ${metadata.platform}; ${metadata.cpu}; ExcelJS ${metadata.exceljs}.\n\nMedian of ${iterations} samples after one warmup for each case. Synthetic general ledger: eight columns, strings with leading zeros, dates, amounts, cached formulas, widths, frozen header, number formats and print settings. Export includes workbook construction and serialization. Import uses the same compressed ExcelJS output for both readers. Output checks run outside timings. MinkExcel reads values; ExcelJS also builds its richer presentation model, so these are different feature sets.\n\n| Data rows | MinkExcel export ms | ExcelJS export ms | MinkExcel import ms | ExcelJS import ms | MinkExcel KiB | ExcelJS KiB |\n| ---: | ---: | ---: | ---: | ---: | ---: | ---: |\n${lines.join("\n")}\n\nBoth writers use DEFLATE. MinkExcel uses the platform CompressionStream API; ExcelJS uses its bundled ZIP compressor. Compression settings and workbook XML differ. These are local Bun measurements, not browser measurements or performance guarantees.\n\nRun from the repository root: \`bun run bench\`. Install optional references with \`bun install --cwd tools\` first. Set \`XLSX_BENCH_ITERATIONS\` to change the sample count.\n`;
await Bun.write(
  new URL("results.json", import.meta.url),
  JSON.stringify({ metadata, rows }, null, 2) + "\n",
);
console.log(report);
