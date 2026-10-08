import { expect, test } from 'bun:test';
import { Workbook, readWorkbook, writeWorkbook, type CellValue } from '../src/index.js';
import { child, parseXml } from '../src/xml.js';
import { writeZip } from '../src/zip.js';
import { ZipArchive } from '../src/zip.js';
import { unzip } from './helpers.js';

async function template(date1904 = false) {
	const book = new Workbook(), sheet = book.addWorksheet('Paywise');
	book.creator = 'Original author';
	book.lastModifiedBy = 'Original modifier';
	book.created = book.modified = new Date('2020-01-01T00:00:00Z');
	sheet.addRow(['Name', 'Account', 'Amount', 'Total', new Date('2026-10-08T00:00:00Z')]);
	sheet.addRow(['Old name', '00001234', 10, { formula: 'SUM(C2:C100)', result: 10 }]);
	sheet.getRow(1).font = { bold: true };
	sheet.getRow(1).height = 24;
	sheet.getColumn(1).width = 30;
	sheet.getColumn(2).numFmt = '@';
	sheet.getColumn(3).numFmt = '#,##0.00';
	sheet.mergeCells('A4:B4');
	sheet.getCell('A4').value = 'Merged footer';
	sheet.getCell('C6').value = 99;
	book.addWorksheet('Second');
	const parts = unzip(await writeWorkbook(book));
	parts.set('xl/worksheets/sheet1.xml', parts.get('xl/worksheets/sheet1.xml')!.replace('<pageMargins', '<sheetProtection sheet="1" password="A05B"/><pageMargins'));
	parts.set('xl/comments1.xml', '<comments xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><authors><author>Original</author></authors><commentList><comment ref="B1" authorId="0"><text><t>Keep this comment</t></text></comment></commentList></comments>');
	parts.set('xl/worksheets/_rels/sheet1.xml.rels', '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="comments" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="../comments1.xml"/></Relationships>');
	if (date1904) {
		parts.set('xl/workbook.xml', parts.get('xl/workbook.xml')!.replace('<bookViews>', '<workbookPr date1904="1"/><bookViews>'));
		parts.set('xl/worksheets/sheet1.xml', parts.get('xl/worksheets/sheet1.xml')!.replace('<v>46303</v>', '<v>44841</v>'));
	}
	return { bytes: await writeZip(parts), parts };
}

test('template edits retain styles, widths, heights, merges, protection, comments and original formulas', async () => {
	const { bytes, parts } = await template();
	const book = await readWorkbook(bytes, { preserveTemplate: true });
	const sheet = book.getWorksheet(1)!;
	sheet.getCell('A2').value = 'New & <name> _x0041_';
	sheet.getCell('B2').value = '00000999';
	sheet.getCell('C2').value = 123.45;
	sheet.getCell('C6').value = null;
	sheet.getCell('A8').value = 'Appended';
	sheet.getCell('B8').value = '00000888';
	sheet.getCell('C8').value = 42;
	const output = await writeWorkbook(book), actualParts = unzip(output);
	for (const part of ['xl/styles.xml', 'xl/sharedStrings.xml', 'xl/comments1.xml', 'xl/worksheets/_rels/sheet1.xml.rels'])
		expect(actualParts.get(part)).toBe(parts.get(part));
	const original = parseXml(parts.get('xl/worksheets/sheet1.xml')!), actual = parseXml(actualParts.get('xl/worksheets/sheet1.xml')!);
	for (const name of ['cols', 'mergeCells', 'sheetProtection', 'sheetViews'])
		expect(child(actual, name)).toEqual(child(original, name));
	expect(child(actual, 'sheetData')!.children[0].attributes.ht).toBe('24');
	const roundtrip = (await readWorkbook(output)).worksheets[0];
	expect(roundtrip.getRow(2).values.slice(1)).toEqual(['New & <name> _x0041_', '00000999', 123.45, { formula: 'SUM(C2:C100)', result: 10 }]);
	expect(roundtrip.getCell('C6').value).toBeNull();
	expect(roundtrip.getRow(8).values.slice(1)).toEqual(['Appended', '00000888', 42]);
	expect(child(actual, 'dimension')!.attributes.ref).toBe('A1:E8');
});

test('template no-op export preserves worksheet and arbitrary archive parts byte for byte', async () => {
	const { bytes, parts } = await template();
	const output = unzip(await writeWorkbook(await readWorkbook(bytes, { preserveTemplate: true })));
	for (const [name, original] of parts) expect(output.get(name)).toBe(original);
});

test('template mode imports metadata and allows changing it without removing other properties', async () => {
	const { bytes } = await template();
	const book = await readWorkbook(bytes, { preserveTemplate: true });
	expect(book.creator).toBe('Original author');
	expect(book.lastModifiedBy).toBe('Original modifier');
	expect(book.created).toEqual(new Date('2020-01-01T00:00:00Z'));
	book.creator = 'PCSTI ERP';
	book.lastModifiedBy = 'Payroll user';
	book.modified = new Date('2026-10-08T12:00:00Z');
	const output = unzip(await writeWorkbook(book));
	expect(output.get('docProps/core.xml')).toContain('PCSTI ERP');
	expect(output.get('docProps/core.xml')).toContain('Payroll user');
	expect(output.get('docProps/core.xml')).toContain('2026-10-08T12:00:00.000Z');
});

test('editing one template metadata field does not require unrelated optional fields', async () => {
	const { parts } = await template();
	parts.set('docProps/core.xml', '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/"><dc:creator>Original author</dc:creator><cp:lastModifiedBy>Original modifier</cp:lastModifiedBy><dc:title>Keep this title</dc:title></cp:coreProperties>');
	const book = await readWorkbook(await writeZip(parts), { preserveTemplate: true });
	book.creator = 'New author';
	const output = unzip(await writeWorkbook(book));
	expect(output.get('docProps/core.xml')).toContain('<dc:creator>New author</dc:creator>');
	expect(output.get('docProps/core.xml')).toContain('<cp:lastModifiedBy>Original modifier</cp:lastModifiedBy>');
	expect(output.get('docProps/core.xml')).toContain('<dc:title>Keep this title</dc:title>');
	book.modified = new Date('2026-10-08T12:00:00Z');
	await expect(writeWorkbook(book)).rejects.toThrow('no editable modified metadata field');
});

test('template replacements round-trip scalar values and every formula cache type', async () => {
	const { bytes } = await template();
	const book = await readWorkbook(bytes, { preserveTemplate: true });
	const values: CellValue[] = [
		false, 0, { error: '#N/A' },
		{ formula: '"text"', result: '_x0041_\r & text' },
		{ formula: '1+1', result: 2 },
		{ formula: '1=2', result: false },
		{ formula: 'NA()', result: { error: '#N/A' } },
		{ formula: 'SUM(C2:C10)' }
	];
	for (const [index, value] of values.entries()) book.worksheets[0].getCell(index + 10, 3).value = value;
	const output = await readWorkbook(await writeWorkbook(book));
	for (const [index, value] of values.entries()) expect(output.worksheets[0].getCell(index + 10, 3).value).toEqual(value);
});

test('template mode rejects unsupported structural and presentation edits', async () => {
	const { bytes } = await template();
	for (const change of [
		(book: Workbook) => book.addWorksheet('Added'),
		(book: Workbook) => book.worksheets.reverse(),
		(book: Workbook) => { book.worksheets[0].getCell('A2').font.bold = true; },
		(book: Workbook) => { book.worksheets[0].getColumn(1).width = 99; },
		(book: Workbook) => { book.worksheets[0].pageSetup.horizontalCentered = true; }
	]) {
		const book = await readWorkbook(bytes, { preserveTemplate: true });
		change(book);
		await expect(writeWorkbook(book)).rejects.toThrow('Template');
	}
});

for (const date1904 of [false, true]) test(`template date edits retain the ${date1904 ? 1904 : 1900} date system`, async () => {
	const { bytes } = await template(date1904);
	const book = await readWorkbook(bytes, { preserveTemplate: true });
	expect(book.worksheets[0].getCell('E1').value).toEqual(new Date('2026-10-08T00:00:00Z'));
	book.worksheets[0].getCell('E1').value = new Date('2026-11-01T12:00:00Z');
	const output = await writeWorkbook(book);
	expect((await readWorkbook(output)).worksheets[0].getCell('E1').value).toEqual(new Date('2026-11-01T12:00:00Z'));
	book.worksheets[0].getCell('F8').value = new Date();
	await expect(writeWorkbook(book)).rejects.toThrow('date-formatted');
});

test('template value changes support repeated exports and clearing rows through values', async () => {
	const { bytes } = await template();
	const book = await readWorkbook(bytes, { preserveTemplate: true });
	book.worksheets[0].getRow(2).values = ['Replacement'];
	for (let index = 0; index < 2; index++) {
		const output = await writeWorkbook(book);
		const row = (await readWorkbook(output)).worksheets[0].getRow(2);
		expect(row.getCell(1).value).toBe('Replacement');
		expect(row.getCell(2).value).toBeNull();
		expect(row.getCell(4).value).toBeNull();
	}
});

test('template mode retains binary media without decoding or modifying it', async () => {
	const { parts } = await template();
	const binary = Uint8Array.from([0, 255, 128, 13, 10, 0, 1]);
	const entries = new Map<string, string | Uint8Array>(parts);
	entries.set('xl/media/image1.png', binary);
	const book = await readWorkbook(await writeZip(entries), { preserveTemplate: true });
	book.worksheets[0].getCell('A2').value = 'Updated';
	const archive = new ZipArchive(await writeWorkbook(book), { fileBytes: 5e6, entries: 100, entryBytes: 5e6, totalBytes: 5e6, compressionRatio: 200 });
	expect(await archive.read('xl/media/image1.png')).toEqual(binary);
});

test('template mode rejects non-finite values even when replacing a blank cell', async () => {
	const { bytes } = await template();
	for (const value of [NaN, Infinity, new Date(NaN)]) {
		const book = await readWorkbook(bytes, { preserveTemplate: true });
		book.worksheets[0].getCell('B4').value = value;
		await expect(writeWorkbook(book)).rejects.toThrow('finite');
	}
});

test('template mode preserves prefixed XML and ignores misleading tags in comments', async () => {
	const { parts } = await template();
	const original = parts.get('xl/worksheets/sheet1.xml')!
		.replace('xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"', 'xmlns:s="http://schemas.openxmlformats.org/spreadsheetml/2006/main"')
		.replace(/<(\/?)([A-Za-z][\w]*)(?=[\s/>])/g, '<$1s:$2')
		.replace('<s:sheetData>', '<!-- <s:sheetData><s:row r="999"/></s:sheetData> --><s:sheetData>');
	parts.set('xl/worksheets/sheet1.xml', original);
	const book = await readWorkbook(await writeZip(parts), { preserveTemplate: true });
	book.worksheets[0].getCell('A2').value = 'New';
	book.worksheets[0].getCell('A8').value = 'Appended';
	const output = await writeWorkbook(book);
	expect(unzip(output).get('xl/worksheets/sheet1.xml')).toContain('<!-- <s:sheetData><s:row r="999"/></s:sheetData> -->');
	expect((await readWorkbook(output)).worksheets[0].getCell('A8').value).toBe('Appended');
});

test('template mode rejects value edits inside array formulas and merged members', async () => {
	const { parts } = await template();
	parts.set('xl/worksheets/sheet1.xml', parts.get('xl/worksheets/sheet1.xml')!.replace('<f>SUM(C2:C100)</f>', '<f t="array" ref="D2:E3">SUM(C2:C100)</f>'));
	const bytes = await writeZip(parts);
	for (const address of ['D2', 'E3', 'B4']) {
		const book = await readWorkbook(bytes, { preserveTemplate: true });
		book.worksheets[0].getCell(address).value = 'Invalid edit';
		await expect(writeWorkbook(book)).rejects.toThrow('Template');
	}
});

test('template mode checks CRCs of retained parts that value imports do not consume', async () => {
	const { bytes } = await template();
	const corrupted = bytes.slice(), view = new DataView(corrupted.buffer);
	let cursor = view.getUint32(corrupted.length - 6, true);
	for (let index = 0; index < view.getUint16(corrupted.length - 12, true); index++) {
		const length = view.getUint16(cursor + 28, true);
		const name = new TextDecoder().decode(corrupted.subarray(cursor + 46, cursor + 46 + length));
		if (name === 'xl/comments1.xml') {
			const crc = (view.getUint32(cursor + 16, true) ^ 1) >>> 0;
			view.setUint32(cursor + 16, crc, true);
			view.setUint32(view.getUint32(cursor + 42, true) + 14, crc, true);
			break;
		}
		cursor += 46 + length + view.getUint16(cursor + 30, true) + view.getUint16(cursor + 32, true);
	}
	await expect(readWorkbook(corrupted, { preserveTemplate: true })).rejects.toThrow('checksum');
	await expect(readWorkbook(corrupted)).resolves.toBeInstanceOf(Workbook);
});
