import { describe, expect, test } from 'bun:test';
import { Workbook, readWorkbook, writeWorkbook } from '../src/index.js';
import { writeZip, crc32 } from '../src/zip.js';
import { fixture, unzip } from './helpers.js';

async function modified(path: string, xml?: string) {
	const parts = unzip(await fixture());
	if (xml === undefined) parts.delete(path);
	else parts.set(path, xml);
	return writeZip(parts);
}
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
			'',
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
			'',
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
	test('rejects invalid addresses, duplicate names and unsupported cell values', () => {
		const book = new Workbook(),
			sheet = book.addWorksheet('Data');
		for (const address of ['A0', 'XFE1', 'A1048577'])
			expect(() => sheet.getCell(address)).toThrow();
		expect(() => book.addWorksheet('data')).toThrow();
		expect(() => sheet.addRow([{}])).toThrow();
	});
	test('uses the standard ZIP CRC32 checksum', () =>
		expect(crc32(new TextEncoder().encode('123456789'))).toBe(0xcbf43926));
	test('round-trips whitespace, XML characters and Excel escapes', async () => {
		const book = new Workbook();
		book
			.addWorksheet('Data')
			.addRow(['  α & <β>  ', '_x0041_', 'line\nnext', '\u0001']);
		expect(
			(await readWorkbook(await writeWorkbook(book))).worksheets[0]
				.getRow(1)
				.values.slice(1)
		).toEqual(['  α & <β>  ', '_x0041_', 'line\nnext', '\u0001']);
	});
});
describe('Workbook import', () => {
	for (const date1904 of [false, true])
		test(`reads independent ExcelJS workbook with date1904=${date1904}`, async () => {
			const sheet = (await readWorkbook(await fixture(date1904))).getWorksheet(
				'Customers'
			)!;
			expect(sheet.rowCount).toBe(3);
			expect(sheet.getCell('A3').value).toBe('001234567');
			expect(sheet.getCell('B3').text).toBe('Ana & Co');
			expect(sheet.getCell('C3').value).toEqual({ formula: '1+2', result: 3 });
			expect(sheet.getCell('D3').value).toBe(false);
			expect(sheet.getCell('E3').text).toBe('Website');
			expect(sheet.getCell('F3').value).toEqual(
				new Date('2026-10-08T12:30:00Z')
			);
		});
	test('rejects malformed XML, DTDs and missing workbook parts', async () => {
		for (const xml of [
			'<worksheet><broken></worksheet>',
			'<!DOCTYPE worksheet [<!ENTITY data "bad">]><worksheet>&data;</worksheet>'
		])
			await expect(
				readWorkbook(await modified('xl/worksheets/sheet1.xml', xml))
			).rejects.toThrow();
		await expect(
			readWorkbook(await modified('xl/workbook.xml'))
		).rejects.toThrow('Missing');
	});
	test('rejects a shared-string part with the wrong root element', async () => {
		const parts = unzip(await fixture());
		parts.set('xl/sharedStrings.xml', parts.get('xl/sharedStrings.xml')!
			.replace('<sst ', '<other ').replace('</sst>', '</other>'));
		await expect(readWorkbook(await writeZip(parts))).rejects.toThrow('shared-string');
	});
	test('enforces archive byte, entry count, compression and decompressed limits', async () => {
		const bytes = await fixture();
		for (const limits of [
			{ fileBytes: 1 },
			{ entries: 1 },
			{ entryBytes: 100 },
			{ totalBytes: 100 },
			{ compressionRatio: 1 }
		])
			await expect(readWorkbook(bytes, limits)).rejects.toThrow();
	});
	test('rejects invalid caller limits instead of silently disabling bounds', async () => {
		const bytes = await fixture();
		for (const key of ['fileBytes', 'entries', 'entryBytes', 'totalBytes', 'rows', 'columns', 'cells'] as const) {
			for (const value of [NaN, Infinity, -1, 1.5])
				await expect(readWorkbook(bytes, { [key]: value })).rejects.toThrow('limit');
		}
		for (const value of [NaN, Infinity, -1, 0])
			await expect(readWorkbook(bytes, { compressionRatio: value })).rejects.toThrow('limit');
	});
	test('undefined optional limits retain defaults and custom byte errors report the actual bound', async () => {
		const bytes = await fixture();
		expect((await readWorkbook(bytes, { rows: undefined })).worksheets[0].rowCount).toBe(3);
		await expect(readWorkbook(bytes, { fileBytes: 100 })).rejects.toThrow('100 bytes');
	});
	test('rejects external worksheets and invalid shared string indices', async () => {
		await expect(
			readWorkbook(
				await modified(
					'xl/worksheets/sheet1.xml',
					'<worksheet><sheetData><row r="1"><c r="A1" t="s"><v>99999</v></c></row></sheetData></worksheet>'
				)
			)
		).rejects.toThrow('shared-string');
		await expect(
			readWorkbook(
				await modified(
					'xl/_rels/workbook.xml.rels',
					'<Relationships><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="https://example.com/sheet.xml" TargetMode="External"/></Relationships>'
				)
			)
		).rejects.toThrow('worksheet relationship');
	});
	test('reads prefixed inline strings and relative relationship paths', async () => {
		const parts = unzip(await fixture());
		parts.set(
			'xl/worksheets/sheet1.xml',
			'<s:worksheet xmlns:s="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><s:sheetData><s:row r="3"><s:c r="C3" t="inlineStr"><s:is><s:t xml:space="preserve">  Ana &amp; Co  </s:t></s:is></s:c></s:row></s:sheetData></s:worksheet>'
		);
		parts.set(
			'xl/_rels/workbook.xml.rels',
			parts
				.get('xl/_rels/workbook.xml.rels')!
				.replace(
					'Target="worksheets/sheet1.xml"',
					'Target="../xl/worksheets/sheet1.xml"'
				)
		);
		expect(
			(await readWorkbook(await writeZip(parts))).worksheets[0].getCell('C3')
				.value
		).toBe('  Ana & Co  ');
	});
	test('rejects duplicate worksheet data sections in both import paths', async () => {
		const bytes = await modified('xl/worksheets/sheet1.xml', '<worksheet><sheetData/><sheetData/></worksheet>');
		for (const signal of [undefined, new AbortController().signal])
			await expect(readWorkbook(bytes, {}, signal)).rejects.toThrow('worksheet data');
	});
	test('ignores rows outside the direct worksheet data section in both import paths', async () => {
		const bytes = await modified('xl/worksheets/sheet1.xml', '<worksheet><ignored><sheetData><row r="1"><c r="A1"><v>999</v></c></row></sheetData></ignored><sheetData><row r="1"><c r="A1"><v>123</v></c></row></sheetData></worksheet>');
		for (const signal of [undefined, new AbortController().signal])
			expect((await readWorkbook(bytes, {}, signal)).worksheets[0].getCell('A1').value).toBe(123);
	});
	for (const [address, reason, row] of [
		['A10002', '10000-row', 10002],
		['CW1', '100-column', 1],
		['Z5000', '100000-cell', 5000]
	] as const)
		test(`rejects sparse oversized ${address}`, async () => {
			await expect(
				readWorkbook(
					await modified(
						'xl/worksheets/sheet1.xml',
						`<worksheet><sheetData><row r="${row}"><c r="${address}" t="inlineStr"><is><t>Value</t></is></c></row></sheetData></worksheet>`
					)
				)
			).rejects.toThrow(reason);
		});
	test('honors cancellation for public read and write calls', async () => {
		const controller = new AbortController();
		controller.abort();
		await expect(
			readWorkbook(await fixture(), {}, controller.signal)
		).rejects.toThrow();
		const book = new Workbook();
		book.addWorksheet('Data');
		await expect(writeWorkbook(book, controller.signal)).rejects.toThrow();
	});
});

test('preserves carriage returns and cached formula error text', async () => {
	const book = new Workbook();
	book
		.addWorksheet('Data')
		.addRow([
			'first\rsecond\r\nthird',
			{ formula: '1/0', result: { error: '#DIV/0!' } },
			{ error: '#VALUE!' }
		]);
	const actual = (await readWorkbook(await writeWorkbook(book))).worksheets[0];
	expect(actual.getCell('A1').value).toBe('first\rsecond\r\nthird');
	expect(actual.getCell('B1').text).toBe('#DIV/0!');
	expect(actual.getCell('C1').text).toBe('#VALUE!');
});
