import ExcelJS from '../tools/node_modules/exceljs/excel.js';
import { readWorkbook, writeWorkbook } from '../dist/index.js';
import { payrollTemplate, editPayroll, verifyPayroll } from './template-fixture.ts';
import { cpus } from 'node:os';
import { readFile, writeFile } from 'node:fs/promises';

const iterations = Number(process.env.XLSX_BENCH_ITERATIONS ?? 7);
if (!Number.isInteger(iterations) || iterations < 1) throw new Error('XLSX_BENCH_ITERATIONS must be a positive integer.');
async function median<T>(run: () => Promise<T>) {
	await run();
	const samples: number[] = [];
	let result!: T;
	for (let index = 0; index < iterations; index++) {
		const start = performance.now();
		result = await run();
		samples.push(performance.now() - start);
	}
	samples.sort((a, b) => a - b);
	return { ms: samples[Math.floor(samples.length / 2)], samples, result };
}
const rows = [];
for (const count of [100, 1_000, 10_000]) {
	const bytes = await payrollTemplate(count);
	const limits = { rows: count + 1, cells: (count + 2) * 4 };
	const values = await median(() => readWorkbook(bytes, limits));
	const preserved = await median(() => readWorkbook(bytes, { ...limits, preserveTemplate: true }));
	const excel = await median(async () => {
		const book = new ExcelJS.Workbook();
		await book.xlsx.load(bytes as unknown as ExcelJS.Buffer);
		return book;
	});
	const minkEdit = await median(async () => {
		const book = await readWorkbook(bytes, { ...limits, preserveTemplate: true });
		editPayroll(book, count);
		return writeWorkbook(book);
	});
	const excelEdit = await median(async () => {
		const book = new ExcelJS.Workbook();
		await book.xlsx.load(bytes as unknown as ExcelJS.Buffer);
		editPayroll(book, count);
		return new Uint8Array(await book.xlsx.writeBuffer());
	});
	// Independent validation of every edited value and the template features, outside timings.
	await verifyPayroll(minkEdit.result, count);
	await verifyPayroll(excelEdit.result, count);
	rows.push({ rows: count, minkexcelValueImportMs: values.ms, minkexcelTemplateImportMs: preserved.ms,
		exceljsImportMs: excel.ms, minkexcelEditMs: minkEdit.ms, exceljsEditMs: excelEdit.ms,
		inputBytes: bytes.length, minkexcelBytes: minkEdit.result.length, exceljsBytes: excelEdit.result.length,
		samples: { valueImport: values.samples, templateImport: preserved.samples, exceljsImport: excel.samples, minkexcelEdit: minkEdit.samples, exceljsEdit: excelEdit.samples } });
}
const isBun = Boolean(process.versions.bun);
const jsonName = isBun ? 'template.json' : 'template-node.json';
const metadata = { date: new Date().toISOString(), runtime: isBun ? `Bun ${process.versions.bun}` : `Node ${process.versions.node}`,
	platform: `${process.platform} ${process.arch}`, cpu: cpus()[0]?.model, minkexcel: JSON.parse(await readFile(new URL('../package.json', import.meta.url), 'utf8')).version,
	exceljs: JSON.parse(await readFile(new URL('../tools/node_modules/exceljs/package.json', import.meta.url), 'utf8')).version, iterations, warmups: 1 };
const report = `# Payroll template editing benchmark

${metadata.date}; ${metadata.runtime}; ${metadata.platform}; ${metadata.cpu}; MinkExcel ${metadata.minkexcel}; ExcelJS ${metadata.exceljs}.

Median of ${iterations} samples after one warmup. The independent ExcelJS producer creates a synthetic payroll template with bank account strings, amounts, column formats, styled headers with borders and fills, merges, row height, frozen headings, a comment, sheet protection, print centering, a total formula and an instructions sheet. Both readers receive the same input bytes. Fixture construction is outside timings. The edit workflow includes import, updating every employee's first three columns and export. Every edited value and the listed template features are checked with ExcelJS outside timings. MinkExcel preserves opaque archive parts for cell value editing; it does not expose a complete editable presentation model. The value-only import is included to show the cost of retaining a template and does not retain its formatting. No AbortSignal is supplied. Neither library calculates formula caches.

| Employees | MinkExcel values import ms | MinkExcel template import ms | ExcelJS import ms | MinkExcel edit workflow ms | ExcelJS edit workflow ms |
| ---: | ---: | ---: | ---: | ---: | ---: |
${rows.map(row => `| ${row.rows.toLocaleString('en-US')} | ${row.minkexcelValueImportMs.toFixed(1)} | ${row.minkexcelTemplateImportMs.toFixed(1)} | ${row.exceljsImportMs.toFixed(1)} | ${row.minkexcelEditMs.toFixed(1)} | ${row.exceljsEditMs.toFixed(1)} |`).join('\n')}

Raw samples and input/output byte sizes are in ${jsonName}. These local in-memory measurements do not measure browser performance or peak memory. Reproduce with \`${isBun ? 'bun run bench:template' : 'bun run bench:template:node'}\` after building the package and installing tools.
`;
await writeFile(new URL(jsonName, import.meta.url), JSON.stringify({ metadata, rows }, null, 2) + '\n');
console.log(report);
