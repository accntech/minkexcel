import ExcelJS from '../tools/node_modules/exceljs/excel.js';
import { strict as assert } from 'node:assert';

export const templateDate = new Date('2026-10-08T00:00:00Z');

/** Synthetic payroll template; no private application data or assets. */
export async function payrollTemplate(count: number): Promise<Uint8Array> {
	const book = new ExcelJS.Workbook();
	book.creator = 'Template author';
	book.lastModifiedBy = 'Template editor';
	book.created = book.modified = templateDate;
	const sheet = book.addWorksheet('Paywise', { views: [{ state: 'frozen', ySplit: 2 }] });
	sheet.columns = [{ width: 30 }, { width: 22 }, { width: 16 }, { width: 16 }];
	sheet.getColumn(2).numFmt = '@';
	sheet.getColumn(3).numFmt = '#,##0.00';
	sheet.getColumn(4).numFmt = '#,##0.00';
	sheet.addRow(['PAYROLL EXPORT']);
	sheet.mergeCells('A1:C1');
	sheet.getRow(1).height = 24;
	sheet.getCell('A1').font = { bold: true, name: 'Arial', size: 14 };
	const header = sheet.addRow(['Name', 'Bank Account', 'Amount', 'Total']);
	header.font = { bold: true };
	header.eachCell((cell) => {
		cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE7E6E6' } };
		cell.border = { bottom: { style: 'double', color: { argb: 'FF000000' } } };
	});
	sheet.getCell('B2').note = 'Keep leading zeros in bank account numbers.';
	for (let row = 1; row <= count; row++) sheet.addRow([`Old employee ${row}`, String(row).padStart(12, '0'), row / 100]);
	sheet.getCell('D3').value = { formula: `SUM(C3:C${count + 2})`, result: count * (count + 1) / 200 };
	sheet.pageSetup = { orientation: 'landscape', horizontalCentered: true, paperSize: 9 };
	await sheet.protect('template', { spinCount: 1 });
	book.addWorksheet('Instructions').addRow(['Change the first three columns only.']);
	return new Uint8Array(await book.xlsx.writeBuffer());
}

export function editPayroll(book: { worksheets: Array<{ getCell(row: number, column: number): { value: unknown } }> }, count: number): void {
	for (let row = 1; row <= count; row++) {
		book.worksheets[0].getCell(row + 2, 1).value = `Employee ${row}`;
		book.worksheets[0].getCell(row + 2, 2).value = String(row + 100).padStart(12, '0');
		book.worksheets[0].getCell(row + 2, 3).value = row / 10;
	}
}

export async function verifyPayroll(bytes: Uint8Array, count: number): Promise<void> {
	const book = new ExcelJS.Workbook();
	await book.xlsx.load(bytes as unknown as ExcelJS.Buffer);
	const sheet = book.getWorksheet('Paywise')!;
	for (let row = 1; row <= count; row++) {
		assert.equal(sheet.getCell(row + 2, 1).value, `Employee ${row}`);
		assert.equal(sheet.getCell(row + 2, 2).value, String(row + 100).padStart(12, '0'));
		assert.equal(sheet.getCell(row + 2, 3).value, row / 10);
	}
	assert.equal(sheet.getColumn(1).width, 30);
	assert.equal(sheet.getColumn(2).numFmt, '@');
	assert.equal(sheet.getColumn(3).numFmt, '#,##0.00');
	assert.equal(sheet.getRow(1).height, 24);
	assert.equal(sheet.getCell('B1').value, 'PAYROLL EXPORT');
	assert.equal(sheet.getCell('A1').font.bold, true);
	assert.equal(sheet.getCell('A2').border.bottom?.style, 'double');
	assert.equal(sheet.getCell('A2').fill.type, 'pattern');
	assert.equal(sheet.getCell('B2').note, 'Keep leading zeros in bank account numbers.');
	assert.equal(sheet.sheetProtection.sheet, true);
	assert.equal(sheet.pageSetup.horizontalCentered, true);
	assert.equal(sheet.views[0].ySplit, 2);
	assert.equal((sheet.getCell('D3').value as ExcelJS.CellFormulaValue).formula, `SUM(C3:C${count + 2})`);
	assert.equal(book.worksheets[1].getCell('A1').value, 'Change the first three columns only.');
}
