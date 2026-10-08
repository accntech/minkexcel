import ExcelJS from '../tools/node_modules/exceljs/excel.js';
import { Workbook, readWorkbook, writeWorkbook, type CellValue } from '../dist/index.js';
import { strict as assert } from 'node:assert';
import { cpus } from 'node:os';
import { readFile, writeFile } from 'node:fs/promises';

type Workload = 'numeric' | 'text' | 'mixed';
const iterations = Number(process.env.XLSX_BENCH_ITERATIONS ?? 7);
if (!Number.isInteger(iterations) || iterations < 1)
	throw new Error('XLSX_BENCH_ITERATIONS must be a positive integer.');
const date = new Date('2026-10-08T00:00:00Z');
const headers = Array.from({ length: 8 }, (_, index) => `Field ${index + 1}`);

function values(workload: Workload, row: number): CellValue[] {
	if (workload === 'numeric')
		return [row, row / 10, -row, 0, row % 100, row * 2, row / 3, row % 2];
	if (workload === 'text')
		return [String(row).padStart(9, '0'), `Item ${row % 200}`, 'α & <Co> 🧾',
			'=1+1', 'literal_text', '  spaces  ', `line\n${row}`, '"quoted"'];
	return [String(row).padStart(9, '0'), `Item ${row}`, date, row % 2 === 0,
		row / 10, { formula: `E${row + 1}*2`, result: row / 10 * 2 }, null,
		{ error: '#N/A' }];
}

function populate(book: Workbook | ExcelJS.Workbook, workload: Workload, count: number) {
	book.creator = 'MinkExcel benchmark';
	book.created = date;
	book.modified = date;
	const sheet = book.addWorksheet('Data', { views: [{ state: 'frozen', ySplit: 1 }] });
	sheet.columns = headers.map(() => ({ width: 20 }));
	sheet.addRow(headers).font = { bold: true };
	for (let row = 1; row <= count; row++) sheet.addRow(values(workload, row));
	if (workload === 'mixed') sheet.getColumn(3).numFmt = 'yyyy-mm-dd';
	if (workload === 'numeric') sheet.getColumn(2).numFmt = '#,##0.00';
	return book;
}

async function median<T>(run: () => Promise<T>) {
	await run();
	const samples: number[] = [];
	let result!: T;
	for (let sample = 0; sample < iterations; sample++) {
		const start = performance.now();
		result = await run();
		samples.push(performance.now() - start);
	}
	samples.sort((a, b) => a - b);
	return { ms: samples[Math.floor(samples.length / 2)], result, samples };
}

const rows = [];
for (const workload of ['numeric', 'text', 'mixed'] as const) {
	for (const count of [100, 1_000, 10_000]) {
		const ownExport = await median(() => writeWorkbook(populate(new Workbook(), workload, count) as Workbook));
		const foreignExport = await median(async () => new Uint8Array(await
			(populate(new ExcelJS.Workbook(), workload, count) as ExcelJS.Workbook).xlsx.writeBuffer()));
		const ownImport = await median(() => readWorkbook(foreignExport.result));
		const foreignImport = await median(async () => {
			const book = new ExcelJS.Workbook();
			await book.xlsx.load(foreignExport.result as unknown as ExcelJS.Buffer);
			return book;
		});
		const independent = new ExcelJS.Workbook();
		await independent.xlsx.load(ownExport.result as unknown as ExcelJS.Buffer);
		const roundtrip = await readWorkbook(ownExport.result);
		// Check every cell, including blanks, with both readers outside the timings.
		for (const book of [ownImport.result, foreignImport.result, independent, roundtrip]) {
			const sheet = book.worksheets[0];
			assert.equal(sheet.rowCount, count + 1);
			for (let row = 1; row <= count; row++) {
				const expected = values(workload, row);
				for (let column = 1; column <= 8; column++)
					assert.deepEqual(sheet.getCell(row + 1, column).value, expected[column - 1]);
			}
		}
		rows.push({ workload, rows: count, minkexcelExportMs: ownExport.ms,
			exceljsExportMs: foreignExport.ms, minkexcelImportMs: ownImport.ms,
			exceljsImportMs: foreignImport.ms, minkexcelBytes: ownExport.result.length,
			exceljsBytes: foreignExport.result.length,
			samples: { minkexcelExport: ownExport.samples, exceljsExport: foreignExport.samples,
				minkexcelImport: ownImport.samples, exceljsImport: foreignImport.samples } });
	}
}
const excelPackage = JSON.parse(await readFile(new URL('../tools/node_modules/exceljs/package.json', import.meta.url), 'utf8'));
const isBun = Boolean(process.versions.bun);
const jsonName = isBun ? 'matrix.json' : 'matrix-node.json';
const reportName = isBun ? 'MATRIX.md' : 'MATRIX-NODE.md';
const command = isBun ? 'bun run bench:matrix' : 'npm run bench:node';
const metadata = { date: new Date().toISOString(), runtime: isBun ? `Bun ${process.versions.bun}` : `Node ${process.versions.node}`,
	platform: `${process.platform} ${process.arch}`, cpu: cpus()[0]?.model,
	exceljs: excelPackage.version, iterations, warmups: 1, columns: 8, abortSignal: false };
const lines = rows.map(row => `| ${row.workload} | ${row.rows.toLocaleString('en-US')} | ${row.minkexcelExportMs.toFixed(1)} | ${row.exceljsExportMs.toFixed(1)} | ${row.minkexcelImportMs.toFixed(1)} | ${row.exceljsImportMs.toFixed(1)} |`);
const report = `# ExcelJS comparison across workloads

${metadata.date}; ${metadata.runtime}; ${metadata.platform}; ${metadata.cpu}; ExcelJS ${metadata.exceljs}.

Median of ${iterations} samples after one warmup. Uses the built package ES modules. Each worksheet has eight columns and a header. Numeric data includes fractions, negatives and zeros; text includes leading-zero identifiers, Unicode, XML characters, literal formula-like text and whitespace; mixed data includes strings, dates, booleans, numbers, cached formulas, blanks and errors. Both writers apply the same supported styles. Export includes workbook construction. Both readers receive the same compressed ExcelJS output. No AbortSignal is supplied. Every data cell is checked with both readers outside timed samples, including a MinkExcel roundtrip. Raw samples and file sizes are in ${jsonName}. Literal Excel escape sequences are tested separately by the reliability comparison because ExcelJS 4.4.0 does not preserve them in these checks.

| Workload | Data rows | MinkExcel export ms | ExcelJS export ms | MinkExcel import ms | ExcelJS import ms |
| --- | ---: | ---: | ---: | ---: | ---: |
${lines.join('\n')}

MinkExcel reads values while ExcelJS builds a richer presentation model. These local document-model results do not compare ExcelJS streaming, browser performance or peak memory, and do not establish a universal speed advantage. Reproduce with \`${command}\` after installing tools and building the package. The Node command requires a runtime with native TypeScript type stripping; it was verified on Node 24.12.0.
`;
await writeFile(new URL(jsonName, import.meta.url), JSON.stringify({ metadata, rows }, null, 2) + '\n');
await writeFile(new URL(reportName, import.meta.url), report);
console.log(report);
