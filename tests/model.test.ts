import { describe, expect, test } from 'bun:test';
import { Workbook, XlsxError, type CellValue } from '../src/index.js';

describe('worksheet model', () => {
	test('address and numeric access identify the same cell at column boundaries', () => {
		const sheet = new Workbook().addWorksheet('Data');
		for (const [column, letter] of [[1, 'A'], [26, 'Z'], [27, 'AA'], [702, 'ZZ'], [703, 'AAA'], [16384, 'XFD']] as const) {
			const cell = sheet.getCell(1048576, column);
			expect(sheet.getCell(`${letter}1048576`)).toBe(cell);
			expect(cell.address).toBe(`${letter}1048576`);
			expect(sheet.getColumn(column).letter).toBe(letter);
		}
	});

	test('rejects invalid addresses and numeric positions', () => {
		const sheet = new Workbook().addWorksheet('Data');
		for (const address of ['', 'a1', '$A$1', 'A0', 'A01', 'XFE1', 'A1048577'])
			expect(() => sheet.getCell(address)).toThrow(XlsxError);
		for (const number of [0, -1, 1.5, NaN, Infinity, 1048577])
			expect(() => sheet.getRow(number)).toThrow(XlsxError);
		for (const number of [0, -1, 1.5, NaN, Infinity, 16385]) {
			expect(() => sheet.getColumn(number)).toThrow(XlsxError);
			expect(() => sheet.getCell(1, number)).toThrow(XlsxError);
		}
	});

	test('counts the furthest created positions and appends after sparse rows', () => {
		const sheet = new Workbook().addWorksheet('Data');
		expect([sheet.rowCount, sheet.columnCount]).toEqual([0, 0]);
		sheet.columns = [{ width: 12 }, { width: 24 }];
		expect(sheet.getColumn(2).width).toBe(24);
		expect(sheet.columnCount).toBe(0);
		const cell = sheet.getCell('D8');
		expect(cell.value).toBeNull();
		expect(sheet.getRow(8)).toBe(cell.row);
		expect([sheet.rowCount, sheet.columnCount, cell.row.cellCount]).toEqual([8, 4, 4]);
		expect(sheet.addRow(['next']).number).toBe(9);
	});

	test('replaces row values with a one-based sparse array without retaining old cells', () => {
		const sheet = new Workbook().addWorksheet('Data');
		const row = sheet.addRow(['old', 'old', 'old', 'old', 'old']);
		row.values = ['new', null, undefined, 0, false, ''];
		expect(row.values.length).toBe(7);
		expect(Object.keys(row.values)).toEqual(['1', '4', '5', '6']);
		expect(row.values[1]).toBe('new');
		expect(row.values.slice(4)).toEqual([0, false, '']);
		expect(row.getCell(2).value).toBeNull();
		expect(row.getCell(3).value).toBeNull();
		row.values = [];
		expect(row.values).toEqual([]);
		expect(row.cellCount).toBe(0);
	});

	test('iterates populated rows and cells in position order, including falsy values', () => {
		const sheet = new Workbook().addWorksheet('Data');
		sheet.getCell('C5').value = false;
		sheet.getCell('A5').value = 0;
		sheet.getCell('B5');
		sheet.getCell('B2').value = '';
		sheet.getRow(3);
		const visited: Array<[number, number, CellValue]> = [];
		sheet.eachRow((row, number) => row.eachCell((cell, column) => {
			visited.push([number, column, cell.value]);
		}));
		expect(visited).toEqual([[2, 2, ''], [5, 1, 0], [5, 3, false]]);
	});

	test('cell text uses cached results and plain values without applying number formats', () => {
		const cell = new Workbook().addWorksheet('Data').getCell('A1');
		cell.numFmt = '0.00';
		for (const [value, text] of [
			[null, ''], [0, '0'], [false, 'false'], ['001', '001'],
			[{ formula: '1+1' }, ''], [{ formula: '1-1', result: 0 }, '0'],
			[{ formula: '1=2', result: false }, 'false'],
			[{ formula: '"text"', result: 'text' }, 'text'],
			[{ formula: '1/0', result: { error: '#DIV/0!' } }, '#DIV/0!'],
			[{ error: '#VALUE!' }, '#VALUE!']
		] as Array<[CellValue, string]>) {
			cell.value = value;
			expect(cell.text).toBe(text);
		}
	});

	test('styles inherit from columns and rows with cell overrides', () => {
		const sheet = new Workbook().addWorksheet('Data');
		const column = sheet.getColumn(1), row = sheet.getRow(1), cell = row.getCell(1);
		expect([column.numFmt, row.numFmt, cell.numFmt]).toEqual(['General', 'General', 'General']);
		column.font.name = 'Calibri';
		column.alignment.horizontal = 'right';
		column.numFmt = '0.00';
		expect(cell.font).toEqual({ name: 'Calibri' });
		expect(cell.numFmt).toBe('0.00');
		row.font = { bold: true };
		row.alignment = { vertical: 'middle' };
		row.numFmt = '#,##0';
		const other = row.getCell(2);
		expect(other.font).toEqual({ bold: true });
		expect(other.alignment).toEqual({ vertical: 'middle' });
		expect(other.numFmt).toBe('#,##0');
		cell.font = { italic: true };
		cell.alignment = { horizontal: 'center' };
		cell.numFmt = '0%';
		expect(cell.resolvedStyle).toEqual({
			font: { italic: true }, alignment: { horizontal: 'center' }, numFmt: '0%'
		});
	});

	test('mutating inherited cell styles does not change rows, columns or sibling cells', () => {
		const sheet = new Workbook().addWorksheet('Data');
		const column = sheet.getColumn(1), row = sheet.getRow(1);
		column.font = { name: 'Calibri' };
		column.alignment = { horizontal: 'right' };
		row.font.bold = true;
		row.alignment.wrapText = true;
		row.getCell(1).font.italic = true;
		row.getCell(1).alignment.horizontal = 'center';
		expect(row.font).toEqual({ bold: true });
		expect(row.alignment).toEqual({ wrapText: true });
		expect(column.font).toEqual({ name: 'Calibri' });
		expect(column.alignment).toEqual({ horizontal: 'right' });
		expect(row.getCell(2).font).toEqual({ bold: true });
		expect(row.getCell(2).alignment).toEqual({ wrapText: true });
	});

	for (const numeric of [false, true]) {
		test(`merged cells share their master's value and inherit its style (${numeric ? 'numeric' : 'address'} range)`, () => {
			const sheet = new Workbook().addWorksheet('Data');
			sheet.getCell('A1').value = 'title';
			sheet.getCell('A1').font = { bold: true };
			if (numeric) sheet.mergeCells(1, 1, 2, 2);
			else sheet.mergeCells('A1:B2');
			expect(sheet.merges).toEqual(['A1:B2']);
			expect(sheet.getCell('B2').value).toBe('title');
			expect(sheet.getCell('B2').font).toEqual({ bold: true });
			sheet.getCell('B2').value = 'updated';
			expect(sheet.getCell('A1').value).toBe('updated');
			expect(sheet.getCell('A2').text).toBe('updated');
		});
	}

	test('rejects reversed, incomplete and overlapping merge ranges', () => {
		const sheet = new Workbook().addWorksheet('Data');
		for (const range of ['A1', 'B2:A1', 'B1:A2'])
			expect(() => sheet.mergeCells(range)).toThrow('Invalid merge');
		sheet.mergeCells('A1:B2');
		for (const range of ['A1:B2', 'B2:C3'])
			expect(() => sheet.mergeCells(range)).toThrow('Overlapping');
	});

	test('validates worksheet names and looks them up by exact name', () => {
		const book = new Workbook();
		const sheet = book.addWorksheet("Owner's Report");
		expect(book.getWorksheet("Owner's Report")).toBe(sheet);
		expect(book.getWorksheet("owner's report")).toBeUndefined();
		expect(book.getWorksheet('missing')).toBeUndefined();
		expect(() => book.addWorksheet("OWNER'S REPORT")).toThrow(XlsxError);
		for (const name of ['', 'a'.repeat(32), 'A/B', 'A\\B', 'A*B', 'A?B', 'A:B', 'A[B', 'A]B', "'Start", "End'"])
			expect(() => book.addWorksheet(name)).toThrow(XlsxError);
		expect(book.addWorksheet('a'.repeat(31)).name).toHaveLength(31);
	});

	test('rejects unsupported row values', () => {
		const sheet = new Workbook().addWorksheet('Data');
		for (const value of [{}, [], () => {}, Symbol('value'), 1n])
			expect(() => sheet.addRow([value])).toThrow('Unsupported cell value');
	});
});
