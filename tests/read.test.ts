import { describe, expect, test } from 'bun:test';
import { readWorkbook, XlsxError, XlsxLimitError } from '../src/index.js';
import { writeZip } from '../src/zip.js';
import { fixture, modifiedFixture, unzip } from './helpers.js';

async function modifiedPart(path: string, xml?: string) {
	return modifiedFixture({ [path]: xml });
}

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
				readWorkbook(await modifiedPart('xl/worksheets/sheet1.xml', xml))
			).rejects.toThrow();
		await expect(
			readWorkbook(await modifiedPart('xl/workbook.xml'))
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
				await modifiedPart(
					'xl/worksheets/sheet1.xml',
					'<worksheet><sheetData><row r="1"><c r="A1" t="s"><v>99999</v></c></row></sheetData></worksheet>'
				)
			)
		).rejects.toThrow('shared-string');
		await expect(
			readWorkbook(
				await modifiedPart(
					'xl/_rels/workbook.xml.rels',
					'<Relationships><Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="https://example.com/sheet.xml" TargetMode="External"/></Relationships>'
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
		const bytes = await modifiedPart('xl/worksheets/sheet1.xml', '<worksheet><sheetData/><sheetData/></worksheet>');
		for (const signal of [undefined, new AbortController().signal])
			await expect(readWorkbook(bytes, {}, signal)).rejects.toThrow('worksheet data');
	});
	test('ignores rows outside the direct worksheet data section in both import paths', async () => {
		const bytes = await modifiedPart('xl/worksheets/sheet1.xml', '<worksheet><ignored><sheetData><row r="1"><c r="A1"><v>999</v></c></row></sheetData></ignored><sheetData><row r="1"><c r="A1"><v>123</v></c></row></sheetData></worksheet>');
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
					await modifiedPart(
						'xl/worksheets/sheet1.xml',
						`<worksheet><sheetData><row r="${row}"><c r="${address}" t="inlineStr"><is><t>Value</t></is></c></row></sheetData></worksheet>`
					)
				)
			).rejects.toThrow(reason);
		});
});

// The visitor-based reader and the cancellable tree-based reader must agree.
for (const withSignal of [false, true]) {
	describe(`worksheet import ${withSignal ? 'with' : 'without'} a signal`, () => {
		const signal = withSignal ? new AbortController().signal : undefined;

		test('infers omitted row/cell addresses and decodes inline rich text and all scalar types', async () => {
			const bytes = await modifiedPart('xl/worksheets/sheet1.xml', `<worksheet><sheetData>
				<row><c t="inlineStr"><is><r><t>  Ana</t></r><r><t> &amp; Co  </t></r></is></c><c t="inlineStr"/><c t="str"><v>_x0041_</v></c><c t="d"><v>2026-10-08T12:30:00Z</v></c><c t="b"><v>0</v></c><c t="b"><v>1</v></c><c t="e"/><c/><c t="n"><v/></c></row>
				<row><c r="C2"><v>12.5</v></c><c><v>-3</v></c></row>
			</sheetData></worksheet>`);
			const sheet = (await readWorkbook(bytes, {}, signal)).worksheets[0];
			expect(sheet.getRow(1).values.slice(1)).toEqual([
				'  Ana & Co  ', '', 'A', new Date('2026-10-08T12:30:00Z'), false, true, { error: '#VALUE!' }
			]);
			expect(sheet.getCell('H1').value).toBeNull();
			expect(sheet.getCell('I1').value).toBeNull();
			expect(sheet.getCell('C2').value).toBe(12.5);
			expect(sheet.getCell('D2').value).toBe(-3);
		});

		for (const [label, cell, message] of [
			['invalid boolean', '<c t="b"><v>2</v></c>', 'boolean'],
			['missing boolean', '<c t="b"/>', 'boolean'],
			['non-finite number', '<c><v>Infinity</v></c>', 'numeric'],
			['non-numeric value', '<c t="n"><v>text</v></c>', 'numeric'],
			['invalid ISO date', '<c t="d"><v>invalid</v></c>', 'date'],
			['overflowing serial date', '<c s="1"><v>1e100</v></c>', 'date'],
			['unsupported type', '<c t="unknown"><v>1</v></c>', 'cell type'],
			['negative string index', '<c t="s"><v>-1</v></c>', 'shared-string'],
			['fractional string index', '<c t="s"><v>0.5</v></c>', 'shared-string'],
			['missing string index', '<c t="s"/>', 'shared-string']
		] as const) {
			test(`rejects ${label}`, async () => {
				const bytes = await modifiedFixture({
					'xl/worksheets/sheet1.xml': `<worksheet><sheetData><row>${cell}</row></sheetData></worksheet>`,
					'xl/styles.xml': '<styleSheet><cellXfs><xf numFmtId="0"/><xf numFmtId="14"/></cellXfs></styleSheet>'
				});
				await expect(readWorkbook(bytes, {}, signal)).rejects.toThrow(message);
			});
		}

		for (const [label, rows, message] of [
			['duplicate rows', '<row r="1"/><row r="1"/>', 'row order'],
			['descending rows', '<row r="2"/><row r="1"/>', 'row order'],
			['fractional rows', '<row r="1.5"/>', 'row'],
			['duplicate cells', '<row><c r="A1"/><c r="A1"/></row>', 'cell order'],
			['descending cells', '<row><c r="B1"/><c r="A1"/></row>', 'cell order'],
			['cell in a different row', '<row><c r="A2"/></row>', 'cell order']
		] as const) {
			test(`rejects ${label}`, async () => {
				const bytes = await modifiedPart('xl/worksheets/sheet1.xml', `<worksheet><sheetData>${rows}</sheetData></worksheet>`);
				await expect(readWorkbook(bytes, {}, signal)).rejects.toThrow(message);
			});
		}

		test('accepts exact data-row, column and rectangular cell limits and rejects one beyond', async () => {
			const bytes = await modifiedPart('xl/worksheets/sheet1.xml', '<worksheet><sheetData><row><c><v>1</v></c><c><v>2</v></c></row><row><c><v>3</v></c><c><v>4</v></c></row></sheetData></worksheet>');
			const limits = { rows: 1, columns: 2, cells: 4 };
			expect((await readWorkbook(bytes, limits, signal)).worksheets[0].getCell('B2').value).toBe(4);
			for (const tighter of [{ rows: 0 }, { columns: 1 }, { cells: 3 }])
				await expect(readWorkbook(bytes, { ...limits, ...tighter }, signal)).rejects.toThrow(XlsxLimitError);
		});

		test('recognizes custom date formats without mistaking quoted, escaped or bracketed literals for dates', async () => {
			const bytes = await modifiedFixture({
				'xl/styles.xml': `<styleSheet><numFmts>
					<numFmt numFmtId="164" formatCode="yyyy-mm-dd"/>
					<numFmt numFmtId="165" formatCode="&quot;month&quot; 0.00"/>
					<numFmt numFmtId="166" formatCode="[Red]0.00"/>
					<numFmt numFmtId="167" formatCode="\\d0.00"/>
				</numFmts><cellXfs><xf numFmtId="164"/><xf numFmtId="165"/><xf numFmtId="166"/><xf numFmtId="167"/></cellXfs></styleSheet>`,
				'xl/worksheets/sheet1.xml': '<worksheet><sheetData><row><c s="0"><v>45292.5</v></c><c s="1"><v>45292.5</v></c><c s="2"><v>45292.5</v></c><c s="3"><v>45292.5</v></c><c s="0"><f>DATE(2024,1,1)</f><v>45292</v></c></row></sheetData></worksheet>'
			});
			expect((await readWorkbook(bytes, {}, signal)).worksheets[0].getRow(1).values.slice(1)).toEqual([
				new Date('2024-01-01T12:00:00Z'), 45292.5, 45292.5, 45292.5, { formula: 'DATE(2024,1,1)', result: 45292 }
			]);
		});
	});
}

test('reads workbook parts relocated through package-absolute relationships without optional strings or styles', async () => {
	const parts = unzip(await fixture());
	const workbook = parts.get('xl/workbook.xml')!;
	parts.delete('xl/workbook.xml');
	parts.delete('xl/sharedStrings.xml');
	parts.delete('xl/styles.xml');
	parts.set('_rels/.rels', '<Relationships><Relationship Id="book" Type="urn:test/officeDocument" Target="/reports/book.xml"/></Relationships>');
	parts.set('reports/book.xml', workbook);
	parts.set('reports/_rels/book.xml.rels', '<Relationships><Relationship Id="rId4" Type="urn:test/worksheet" Target="/xl/worksheets/sheet1.xml"/></Relationships>');
	parts.set('xl/worksheets/sheet1.xml', '<worksheet><sheetData><row><c t="inlineStr"><is><t>Relocated</t></is></c></row></sheetData></worksheet>');
	expect((await readWorkbook(await writeZip(parts))).worksheets[0].getCell('A1').value).toBe('Relocated');
});

for (const [target, message] of [
	['../../outside.xml', 'leaves the archive'],
	['https://example.com/sheet.xml', 'relationship target'],
	['worksheets\\sheet1.xml', 'relationship target'],
	['worksheets/sheet1.xml#fragment', 'relationship target'],
	['worksheets/sheet1.xml?query', 'relationship target']
] as const) {
	test(`rejects worksheet relationship target ${target}`, async () => {
		const bytes = await modifiedPart('xl/_rels/workbook.xml.rels', `<Relationships><Relationship Id="rId1" Type="urn:test/worksheet" Target="${target}"/></Relationships>`);
		await expect(readWorkbook(bytes)).rejects.toThrow(message);
	});
}

test('returns typed workbook errors when package documents or required relationships are missing', async () => {
	for (const [path, xml, message] of [
		['_rels/.rels', '<Relationships/>', 'Missing workbook relationship'],
		['_rels/.rels', '<wrong/>', 'Invalid workbook relationships'],
		['xl/workbook.xml', '<wrong/>', 'Invalid workbook document'],
		['xl/workbook.xml', '<workbook/>', 'Missing workbook sheets'],
		['xl/_rels/workbook.xml.rels', '<Relationships/>', 'Missing worksheet relationship'],
		['xl/worksheets/sheet1.xml', undefined, 'Missing workbook part']
	] as const) {
		const pending = readWorkbook(await modifiedPart(path, xml));
		await expect(pending).rejects.toThrow(XlsxError);
		await expect(pending).rejects.toThrow(message);
	}
});
