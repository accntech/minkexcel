import ExcelJS from '../tools/node_modules/exceljs/excel.js';
import { Workbook, readWorkbook, writeWorkbook } from '../dist/index.js';
import { ZipArchive, writeZip } from '../src/zip.js';
import { strict as assert } from 'node:assert';
import { payrollTemplate, editPayroll, verifyPayroll } from './template-fixture.ts';

const limits = { fileBytes: 5 * 1024 * 1024, entries: 1000, entryBytes: 20 * 1024 * 1024,
	totalBytes: 50 * 1024 * 1024, compressionRatio: 200 };
const book = new Workbook();
const literals = ['_x0041_', '_x005F_', 'first\rsecond', '  α & <Co> 🧾  '];
book.addWorksheet('Data').addRow(literals);
const ownBytes = await writeWorkbook(book);
const foreign = new ExcelJS.Workbook();
foreign.addWorksheet('Data').addRow(literals);
const foreignBytes = new Uint8Array(await foreign.xlsx.writeBuffer());
const rows: Array<{ check: string; minkexcel: string; exceljs: string }> = [];

// Independently read the new report export features rather than checking implementation text.
const reportBook = new Workbook();
reportBook.creator = 'Report author';
reportBook.lastModifiedBy = 'Report reviewer';
const reportSheet = reportBook.addWorksheet('Payroll');
reportSheet.addRow(['Employee', 123.45]);
reportSheet.getCell('A1').border = { top: { style: 'thin', color: { argb: 'FF000000' } }, bottom: { style: 'double' } };
reportSheet.pageSetup.horizontalCentered = true;
const independentReport = new ExcelJS.Workbook();
await independentReport.xlsx.load(await writeWorkbook(reportBook) as unknown as ExcelJS.Buffer);
assert.equal(independentReport.lastModifiedBy, 'Report reviewer');
assert.equal(independentReport.worksheets[0].pageSetup.horizontalCentered, true);
assert.equal(independentReport.worksheets[0].getCell('A1').border.top?.style, 'thin');
assert.equal(independentReport.worksheets[0].getCell('A1').border.top?.color?.argb, 'FF000000');
assert.equal(independentReport.worksheets[0].getCell('A1').border.bottom?.style, 'double');
rows.push({ check: 'Report borders, print centering and independent last-modifier metadata', minkexcel: 'Preserved', exceljs: 'Supported' });

const templateBytes = await payrollTemplate(10);
const ownTemplate = await readWorkbook(templateBytes, { preserveTemplate: true });
editPayroll(ownTemplate, 10);
await verifyPayroll(await writeWorkbook(ownTemplate), 10);
const excelTemplate = new ExcelJS.Workbook();
await excelTemplate.xlsx.load(templateBytes as unknown as ExcelJS.Buffer);
editPayroll(excelTemplate, 10);
await verifyPayroll(new Uint8Array(await excelTemplate.xlsx.writeBuffer()), 10);
rows.push({ check: 'Edit payroll template values; retain formatting, comments, protection and layout', minkexcel: 'Preserved (template mode)', exceljs: 'Preserved' });

async function ownValues(bytes: Uint8Array) {
	return (await readWorkbook(bytes)).worksheets[0].getRow(1).values.slice(1);
}
async function excelValues(bytes: Uint8Array) {
	const workbook = new ExcelJS.Workbook();
	await workbook.xlsx.load(bytes as unknown as ExcelJS.Buffer);
	return literals.map((_, index) => workbook.worksheets[0].getCell(1, index + 1).value);
}
// Use both independent readers to check each writer's strings.
const ownWithOwn = await ownValues(ownBytes), ownWithExcel = await excelValues(ownBytes);
const foreignWithOwn = await ownValues(foreignBytes), foreignWithExcel = await excelValues(foreignBytes);
assert.deepEqual(ownWithOwn, literals);
assert.deepEqual(ownWithExcel, literals);
for (let index = 0; index < literals.length; index++) {
	rows.push({ check: `Preserve string ${JSON.stringify(literals[index])} with both readers`,
		minkexcel: ownWithOwn[index] === literals[index] && ownWithExcel[index] === literals[index] ? 'Preserved' : 'Changed',
		exceljs: foreignWithOwn[index] === literals[index] && foreignWithExcel[index] === literals[index] ? 'Preserved' : 'Changed' });
}

const corrupt = ownBytes.slice(), view = new DataView(corrupt.buffer);
let cursor = view.getUint32(corrupt.length - 6, true);
let checksumChanged = false;
for (let index = 0; index < view.getUint16(corrupt.length - 12, true); index++) {
	const length = view.getUint16(cursor + 28, true);
	const name = new TextDecoder().decode(corrupt.subarray(cursor + 46, cursor + 46 + length));
	if (name === 'xl/worksheets/sheet1.xml') {
		const checksum = (view.getUint32(cursor + 16, true) ^ 1) >>> 0;
		view.setUint32(cursor + 16, checksum, true);
		view.setUint32(view.getUint32(cursor + 42, true) + 14, checksum, true);
		checksumChanged = true;
		break;
	}
	cursor += 46 + length + view.getUint16(cursor + 30, true) + view.getUint16(cursor + 32, true);
}
assert.equal(checksumChanged, true);
const archive = new ZipArchive(ownBytes, limits);
const parts = new Map<string, string>();
// Writer emits these package parts; read them through the owned ZIP reader.
for (const name of ['[Content_Types].xml', '_rels/.rels', 'docProps/core.xml', 'docProps/app.xml',
	'xl/workbook.xml', 'xl/_rels/workbook.xml.rels', 'xl/styles.xml', 'xl/sharedStrings.xml', 'xl/worksheets/sheet1.xml'])
	parts.set(name, new TextDecoder().decode(await archive.read(name)));
const malformed = new Map(parts);
malformed.set('xl/worksheets/sheet1.xml', parts.get('xl/worksheets/sheet1.xml')!.replace('<worksheet ', '<worksheet xmlns:xml="urn:wrong" '));
const dtd = new Map(parts);
dtd.set('xl/worksheets/sheet1.xml', parts.get('xl/worksheets/sheet1.xml')!.replace('<worksheet ', '<!DOCTYPE worksheet [<!ENTITY bad "data">]><worksheet '));

async function acceptance(read: () => Promise<unknown>) {
	try { await read(); return 'Accepted'; } catch { return 'Rejected'; }
}
for (const [check, bytes] of [
	['Reject mismatched worksheet CRC32 (both ZIP headers changed consistently)', corrupt],
	['Reject truncated ZIP archive', ownBytes.subarray(0, ownBytes.length - 10)],
	['Reject reserved XML namespace rebinding', await writeZip(malformed)],
	['Reject worksheet DTD', await writeZip(dtd)],
] as const) {
	const minkexcel = await acceptance(() => ownValues(bytes));
	const exceljs = await acceptance(() => excelValues(bytes));
	assert.equal(minkexcel, 'Rejected', check);
	rows.push({ check, minkexcel, exceljs });
}
const excelPackage = await Bun.file(new URL('../tools/node_modules/exceljs/package.json', import.meta.url)).json();
const metadata = { date: new Date().toISOString(), runtime: `Bun ${Bun.version}`, exceljs: excelPackage.version };
const lines = rows.map(row => `| ${row.check.replace(/\|/g, '\\|')} | ${row.minkexcel} | ${row.exceljs} |`);
const report = `# File integrity and string preservation comparison

${metadata.date}; ${metadata.runtime}; ExcelJS ${metadata.exceljs}.

String checks compare each writer's output with both independent readers against the original values. Integrity checks feed identical bytes to both readers with default public import options. The CRC case changes the declared checksum in both ZIP headers, leaving payload bytes unchanged. It tests whether the reader validates the declared checksum against actual data. The DTD case contains an internal declaration; no external resources are used.

| Check | MinkExcel | ExcelJS ${metadata.exceljs} |
| --- | --- | --- |
${lines.join('\n')}

These are specific reproducible checks, not an overall reliability score. ExcelJS supports a broader workbook model; passing these checks does not establish feature parity or reliability for every workbook. Reproduce with \`bun run test:interop\` after installing tools and building the package.
`;
await Bun.write(new URL('reliability.json', import.meta.url), JSON.stringify({ metadata, rows,
	strings: { input: literals, ownWithOwn, ownWithExcel, foreignWithOwn, foreignWithExcel } }, null, 2) + '\n');
console.log(report);
