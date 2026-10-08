import { describe, expect, test } from 'bun:test';
import { Workbook, readWorkbook, writeWorkbook } from '../src/index.js';
import { child, children, parseXml } from '../src/xml.js';
import { unzip } from './helpers.js';

describe('Workbook export', () => {
	test('identifies MinkExcel in workbook metadata and preserves custom creators', async () => {
		const book = new Workbook();
		book.addWorksheet('Data').addRow(['Example']);
		const parts = unzip(await writeWorkbook(book));
		expect(parts.get('docProps/app.xml')).toContain(
			'<Application>MinkExcel</Application>'
		);
		expect(parts.get('docProps/core.xml')).toContain(
			'<dc:creator>MinkExcel</dc:creator>'
		);
		book.creator = 'Libro';
		const customized = unzip(await writeWorkbook(book));
		expect(customized.get('docProps/core.xml')).toContain(
			'<dc:creator>Libro</dc:creator>'
		);
	});
	test('preserves typed scalars, literal strings, identifiers, Unicode, dates and formula caches', async () => {
		const book = new Workbook();
		const sheet = book.addWorksheet('Data & amounts');
		sheet.addRow([
			'001234567',
			'  Ana & <Co> 🧾  ',
			'=1+1',
			'_x0041_',
			'line\nnext\u0001',
			true,
			null
		]);
		sheet.addRow([
			12.34,
			-5,
			0,
			{ formula: 'SUM(A2:C2)', result: 7.34 },
			new Date('2026-10-08T00:00:00Z')
		]);
		const bytes = await writeWorkbook(book),
			parts = unzip(bytes);
		expect(parts.get('xl/sharedStrings.xml')).toContain(
			'<t xml:space="preserve">001234567</t>'
		);
		expect(parts.get('xl/sharedStrings.xml')).toContain(
			'  Ana &amp; &lt;Co&gt; 🧾  '
		);
		expect(parts.get('xl/sharedStrings.xml')).toContain('>_x005F_x0041_</t>');
		expect(parts.get('xl/worksheets/sheet1.xml')).toContain(
			'<c r="F1" t="b"><v>1</v></c>'
		);
		expect(parts.get('xl/worksheets/sheet1.xml')).toContain(
			'<c r="D2"><f>SUM(A2:C2)</f><v>7.34</v></c>'
		);
		const actual = (await readWorkbook(bytes)).worksheets[0];
		expect(actual.getRow(1).values.slice(1)).toEqual([
			'001234567',
			'  Ana & <Co> 🧾  ',
			'=1+1',
			'_x0041_',
			'line\nnext\u0001',
			true
		]);
		expect(actual.getRow(2).values.slice(1)).toEqual([
			12.34,
			-5,
			0,
			{ formula: 'SUM(A2:C2)', result: 7.34 },
			new Date('2026-10-08T00:00:00Z')
		]);
	});
	test('serializes report styling, merges, frozen panes, print settings and metadata', async () => {
		const book = new Workbook();
		book.created = new Date('2026-10-08T00:00:00Z');
		book.modified = book.created;
		const sheet = book.addWorksheet("Owner's Report", {
			views: [{ state: 'frozen', ySplit: 2 }]
		});
		sheet.getCell('A1').value = 'Report';
		sheet.mergeCells('A1:C1');
		const header = sheet.getRow(3);
		header.values = ['Name', 'Amount', 'Date'];
		header.height = 30;
		header.font = { bold: true, size: 13, color: { argb: 'FFFFFFFF' } };
		header.alignment = { vertical: 'middle', wrapText: true };
		header.getCell(1).fill = {
			type: 'pattern',
			pattern: 'solid',
			fgColor: { argb: 'FF334155' }
		};
		sheet.getColumn(1).width = 28;
		sheet.getColumn(1).alignment = { vertical: 'top' };
		sheet.addRow(['Acme', 1.25, new Date('2026-10-08T00:00:00Z')]);
		sheet.getCell('B4').numFmt = '"₱"#,##0.00';
		sheet.getCell('B4').alignment = { horizontal: 'right' };
		sheet.autoFilter = { from: 'A3', to: 'C4' };
		sheet.pageSetup = {
			orientation: 'landscape',
			paperSize: 5,
			fitToPage: true,
			fitToWidth: 1,
			fitToHeight: 0,
			printTitlesRow: '3:3',
			printArea: 'A1:C4',
			margins: {
				left: 0.2,
				right: 0.2,
				top: 0.5,
				bottom: 0.5,
				header: 0.2,
				footer: 0.2
			}
		};
		sheet.headerFooter.oddFooter = '&LLibro &RPage &P of &N';
		const parts = unzip(await writeWorkbook(book)),
			xml = parts.get('xl/worksheets/sheet1.xml')!;
		expect(xml).toContain('<mergeCell ref="A1:C1"/>');
		expect(xml).toContain('ySplit="2"');
		expect(xml).toContain('width="28"');
		expect(xml).toContain('ht="30"');
		expect(xml).toContain('orientation="landscape"');
		expect(xml).toContain('<autoFilter ref="A3:C4"/>');
		expect(xml).toContain('&amp;LLibro &amp;RPage &amp;P of &amp;N');
		expect(parts.get('xl/styles.xml')).toContain('<color rgb="FFFFFFFF"/>');
		expect(parts.get('xl/styles.xml')).toContain('FF334155');
		expect(parts.get('xl/styles.xml')).toContain('&quot;₱&quot;#,##0.00');
		expect(parts.get('xl/workbook.xml')).toContain(
			'&apos;Owner&apos;&apos;s Report&apos;!$A$1:$C$4'
		);
		expect(parts.get('xl/workbook.xml')).toContain(
			'&apos;Owner&apos;&apos;s Report&apos;!$3:$3'
		);
		expect(parts.get('docProps/core.xml')).toContain(
			'2026-10-08T00:00:00.000Z'
		);
	});
	for (const value of [NaN, Infinity, -Infinity])
		test(`rejects ${value}`, async () => {
			const book = new Workbook();
			book.addWorksheet('Data').addRow([value]);
			await expect(writeWorkbook(book)).rejects.toThrow('finite');
		});

});


test('preserves every formula cache type, uncached formulas and literal text', async () => {
	const book = new Workbook();
	book
		.addWorksheet('Data')
		.addRow([
			'first\rsecond\r\nthird',
			{ formula: '1/0', result: { error: '#DIV/0!' } },
			{ error: '#VALUE!' },
			{ formula: '1=2', result: false },
			{ formula: '1=1', result: true },
			{ formula: '"text"', result: '  <text> & _x0041_\r  ' },
			{ formula: '""', result: '' },
			{ formula: '1-1', result: 0 },
			{ formula: 'SUM(A2:A3)' },
			''
		]);
	const bytes = await writeWorkbook(book);
	const actual = (await readWorkbook(bytes)).worksheets[0];
	expect(actual.getRow(1).values).toEqual(book.worksheets[0].getRow(1).values);
	expect(actual.getCell('B1').text).toBe('#DIV/0!');
	expect(actual.getCell('C1').text).toBe('#VALUE!');
	const cells = children(child(parseXml(unzip(bytes).get('xl/worksheets/sheet1.xml')!), 'sheetData')!.children[0], 'c');
	expect(cells.slice(1, 9).map((cell) => cell.attributes.t)).toEqual(['e', 'e', 'b', 'b', 'str', 'str', undefined, undefined]);
	expect(child(cells[8], 'v')).toBeUndefined();
});

test('exports multiple sheets with shared strings deduplicated across sheets', async () => {
	const book = new Workbook();
	book.addWorksheet('First').addRow(['shared', 'first']);
	book.addWorksheet('Second').addRow(['shared', 'second']);
	const bytes = await writeWorkbook(book), parts = unzip(bytes);
	const strings = parseXml(parts.get('xl/sharedStrings.xml')!);
	expect(strings.attributes.uniqueCount).toBe('3');
	expect(children(strings, 'si').map((node) => child(node, 't')!.text)).toEqual(['shared', 'first', 'second']);
	const actual = await readWorkbook(bytes);
	expect(actual.worksheets.map((sheet) => sheet.name)).toEqual(['First', 'Second']);
	expect(actual.getWorksheet('First')!.getRow(1).values.slice(1)).toEqual(['shared', 'first']);
	expect(actual.getWorksheet('Second')!.getRow(1).values.slice(1)).toEqual(['shared', 'second']);
	const types = parseXml(parts.get('[Content_Types].xml')!);
	expect(children(types, 'Override').map((node) => node.attributes.PartName)).toContain('/xl/worksheets/sheet2.xml');
});

test('exports an empty sheet and sparse cells in address order', async () => {
	const book = new Workbook();
	book.addWorksheet('Empty');
	const sheet = book.addWorksheet('Sparse');
	sheet.getCell('AA4').value = 'last';
	sheet.getCell('B4').value = false;
	sheet.getCell('A2').value = 0;
	sheet.autoFilter = { from: { row: 2, column: 1 }, to: { row: 4, column: 27 } };
	const bytes = await writeWorkbook(book), parts = unzip(bytes);
	const empty = parseXml(parts.get('xl/worksheets/sheet1.xml')!);
	expect(child(empty, 'dimension')!.attributes.ref).toBe('A1:A1');
	expect(child(empty, 'sheetData')!.children).toEqual([]);
	const sparse = parseXml(parts.get('xl/worksheets/sheet2.xml')!);
	const rows = children(child(sparse, 'sheetData')!, 'row');
	expect(rows.map((row) => row.attributes.r)).toEqual(['2', '4']);
	expect(children(rows[1], 'c').map((cell) => cell.attributes.r)).toEqual(['B4', 'AA4']);
	expect(child(sparse, 'autoFilter')!.attributes.ref).toBe('A2:AA4');
	const actual = (await readWorkbook(bytes)).getWorksheet('Sparse')!;
	expect(actual.getCell('A2').value).toBe(0);
	expect(actual.getCell('B4').value).toBe(false);
	expect(actual.getCell('AA4').value).toBe('last');
});

test('serializes merged values only at the master and retains member styles', async () => {
	const book = new Workbook(), sheet = book.addWorksheet('Merged');
	sheet.mergeCells('A1:B1');
	sheet.getCell('B1').value = 'Title';
	sheet.getCell('A1').font = { bold: true };
	const parts = unzip(await writeWorkbook(book));
	const row = child(parseXml(parts.get('xl/worksheets/sheet1.xml')!), 'sheetData')!.children[0];
	const [master, member] = children(row, 'c');
	expect(child(master, 'v')!.text).toBe('0');
	expect(member.children).toEqual([]);
	expect(member.attributes.s).toBe(master.attributes.s);
});

test('reuses identical styles and keeps distinct cell overrides', async () => {
	const book = new Workbook(), sheet = book.addWorksheet('Styles');
	sheet.addRow([1, 2, 3]);
	sheet.getRow(1).font = { name: 'Calibri', italic: true };
	sheet.getRow(1).numFmt = '0.00';
	sheet.getCell('C1').numFmt = '0%';
	const parts = unzip(await writeWorkbook(book));
	const cells = child(parseXml(parts.get('xl/worksheets/sheet1.xml')!), 'sheetData')!.children[0].children;
	expect(cells[0].attributes.s).toBe(cells[1].attributes.s);
	expect(cells[2].attributes.s).not.toBe(cells[0].attributes.s);
	const styles = parseXml(parts.get('xl/styles.xml')!);
	expect(child(styles, 'cellXfs')!.children).toHaveLength(3);
	expect(child(styles, 'numFmts')!.children.map((node) => node.attributes.formatCode)).toEqual(['0.00', '0%']);
	expect(child(styles, 'fonts')!.children).toHaveLength(2);
});

test('round-trips dates around the 1900 leap-day discontinuity with fractional days', async () => {
	const dates = ['1899-12-31T00:00:00Z', '1900-02-28T12:00:00Z', '1900-03-01T06:00:00Z'].map((value) => new Date(value));
	const book = new Workbook();
	book.addWorksheet('Dates').addRow(dates);
	const bytes = await writeWorkbook(book);
	const row = child(parseXml(unzip(bytes).get('xl/worksheets/sheet1.xml')!), 'sheetData')!.children[0];
	expect(row.children.map((cell) => child(cell, 'v')!.text)).toEqual(['0', '59.5', '61.25']);
	expect((await readWorkbook(bytes)).worksheets[0].getRow(1).values.slice(1)).toEqual(dates);
});

test('rejects an empty workbook, invalid dates and non-finite formula caches', async () => {
	await expect(writeWorkbook(new Workbook())).rejects.toThrow('at least one worksheet');
	for (const value of [new Date(NaN), { formula: '1/0', result: Infinity }, { formula: '0/0', result: NaN }]) {
		const book = new Workbook();
		book.addWorksheet('Data').addRow([value]);
		await expect(writeWorkbook(book)).rejects.toThrow('finite');
	}
});
