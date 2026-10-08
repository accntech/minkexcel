import { writeZip } from './zip.js';
import { checkpoint } from './async.js';
import {
	Workbook,
	Worksheet,
	Cell,
	XlsxError,
	cellAddress,
	type Style
} from './model.js';
import {
	attributes,
	declaration,
	excelText,
	relationshipNamespace,
	spreadsheetNamespace,
	xml
} from './xml.js';

function numeric(value: number): string {
	if (!Number.isFinite(value))
		throw new XlsxError('Cell numbers must be finite.');
	return String(value);
}
function dateSerial(date: Date): number {
	const value = (date.getTime() - Date.UTC(1899, 11, 31)) / 86_400_000;
	return value >= 60 ? value + 1 : value;
}

class StyleTable {
	private fonts: string[] = [
		'<font><sz val="11"/><name val="Calibri"/></font>'
	];
	private fills: string[] = [
		'<fill><patternFill patternType="none"/></fill>',
		'<fill><patternFill patternType="gray125"/></fill>'
	];
	private formats: string[] = [];
	private records: string[] = [
		'<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>'
	];
	private register(entries: string[], entry: string): number {
		const index = entries.indexOf(entry);
		if (index >= 0) return index;
		entries.push(entry);
		return entries.length - 1;
	}
	id(style: Style): number {
		const font = style.font;
		const fontId =
			font && Object.keys(font).length
				? this.register(
						this.fonts,
						`<font>${font.bold ? '<b/>' : ''}${font.italic ? '<i/>' : ''}${font.size !== undefined ? `<sz val="${xml(font.size)}"/>` : ''}${font.name ? `<name val="${xml(font.name)}"/>` : ''}${font.color ? `<color rgb="${xml(font.color.argb)}"/>` : ''}</font>`
					)
				: 0;
		const fillId = style.fill
			? this.register(
					this.fills,
					`<fill><patternFill patternType="solid"><fgColor rgb="${xml(style.fill.fgColor.argb)}"/><bgColor indexed="64"/></patternFill></fill>`
				)
			: 0;
		const numFmtId =
			style.numFmt && style.numFmt !== 'General'
				? 164 + this.register(this.formats, style.numFmt)
				: 0;
		const alignment =
			style.alignment && Object.keys(style.alignment).length
				? `<alignment${attributes({ ...style.alignment, vertical: style.alignment.vertical === 'middle' ? 'center' : style.alignment.vertical })}/>`
				: '';
		if (!fontId && !fillId && !numFmtId && !alignment) return 0;
		return this.register(
			this.records,
			`<xf${attributes({ numFmtId, fontId, fillId, borderId: 0, xfId: 0, applyFont: fontId ? true : undefined, applyFill: fillId ? true : undefined, applyNumberFormat: numFmtId ? true : undefined, applyAlignment: alignment ? true : undefined })}>${alignment}</xf>`
		);
	}
	toXml(): string {
		return `${declaration}<styleSheet xmlns="${spreadsheetNamespace}"><numFmts count="${this.formats.length}">${this.formats.map((format, index) => `<numFmt numFmtId="${index + 164}" formatCode="${xml(format)}"/>`).join('')}</numFmts><fonts count="${this.fonts.length}">${this.fonts.join('')}</fonts><fills count="${this.fills.length}">${this.fills.join('')}</fills><borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="${this.records.length}">${this.records.join('')}</cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>`;
	}
}

class StringTable {
	readonly values: string[] = [];
	private indices = new Map<string, number>();
	id(value: string): number {
		let index = this.indices.get(value);
		if (index === undefined) {
			index = this.values.length;
			this.values.push(value);
			this.indices.set(value, index);
		}
		return index;
	}
	toXml() {
		return `${declaration}<sst xmlns="${spreadsheetNamespace}" uniqueCount="${this.values.length}">${this.values.map((value) => `<si><t xml:space="preserve">${excelText(value)}</t></si>`).join('')}</sst>`;
	}
}

function cellXml(cell: Cell, styles: StyleTable, strings: StringTable): string {
	const value = cell.master ? null : cell.value;
	const style = cell.resolvedStyle;
	const id = styles.id(
		value instanceof Date && !style.numFmt
			? { ...style, numFmt: 'mm-dd-yy' }
			: style
	);
	const attrs = attributes({ r: cell.address, s: id || undefined });
	if (value === null) return `<c${attrs}/>`;
	if (typeof value === 'string')
		return `<c${attrs} t="s"><v>${strings.id(value)}</v></c>`;
	if (value instanceof Date)
		return `<c${attrs}><v>${numeric(dateSerial(value))}</v></c>`;
	if (typeof value === 'number')
		return `<c${attrs}><v>${numeric(value)}</v></c>`;
	if (typeof value === 'boolean')
		return `<c${attrs} t="b"><v>${Number(value)}</v></c>`;
	if ('error' in value)
		return `<c${attrs} t="e"><v>${xml(value.error)}</v></c>`;
	const result = value.result;
	const type =
		typeof result === 'string'
			? 'str'
			: typeof result === 'boolean'
				? 'b'
				: typeof result === 'object'
					? 'e'
					: undefined;
	const cached =
		result === undefined
			? ''
			: `<v>${typeof result === 'number' ? numeric(result) : typeof result === 'boolean' ? Number(result) : typeof result === 'object' ? xml(result.error) : excelText(result)}</v>`;
	return `<c${attrs}${attributes({ t: type })}><f>${xml(value.formula)}</f>${cached}</c>`;
}

function worksheetXml(
	sheet: Worksheet,
	styles: StyleTable,
	strings: StringTable
): string {
	const setup = sheet.pageSetup;
	const rows = [...sheet.rows.values()].sort((a, b) => a.number - b.number);
	const columns = [...sheet.columnDefinitions.values()].sort(
		(a, b) => a.number - b.number
	);
	const data = rows
		.map((row) => {
			const styleId = styles.id(row.style);
			return `<row${attributes({ r: row.number, ht: row.height, customHeight: row.height !== undefined ? true : undefined, s: styleId || undefined, customFormat: styleId ? true : undefined })}>${[
				...row.cells.values()
			]
				.sort((a, b) => a.column - b.column)
				.map((cell) => cellXml(cell, styles, strings))
				.join('')}</row>`;
		})
		.join('');
	const columnXml = columns
		.map((column) => {
			const styleId = styles.id(column.style);
			return `<col${attributes({ min: column.number, max: column.number, width: column.width, customWidth: column.width !== undefined ? true : undefined, style: styleId || undefined })}/>`;
		})
		.join('');
	const views = sheet.views
		?.map(
			(view) =>
				`<sheetView workbookViewId="0"><pane ySplit="${view.ySplit}" topLeftCell="A${view.ySplit + 1}" activePane="bottomLeft" state="frozen"/><selection pane="bottomLeft" activeCell="A${view.ySplit + 1}" sqref="A${view.ySplit + 1}"/></sheetView>`
		)
		.join('');
	const address = (position: string | { row: number; column: number }) =>
		typeof position === 'string'
			? position
			: cellAddress(position.row, position.column);
	return `${declaration}<worksheet xmlns="${spreadsheetNamespace}" xmlns:r="${relationshipNamespace}">${setup.fitToPage ? '<sheetPr><pageSetUpPr fitToPage="1"/></sheetPr>' : ''}<dimension ref="A1:${cellAddress(Math.max(sheet.rowCount, 1), Math.max(sheet.columnCount, 1))}"/>${views ? `<sheetViews>${views}</sheetViews>` : ''}<sheetFormatPr defaultRowHeight="15"/>${columnXml ? `<cols>${columnXml}</cols>` : ''}<sheetData>${data}</sheetData>${sheet.autoFilter ? `<autoFilter ref="${xml(address(sheet.autoFilter.from) + ':' + address(sheet.autoFilter.to))}"/>` : ''}${sheet.merges.length ? `<mergeCells count="${sheet.merges.length}">${sheet.merges.map((range) => `<mergeCell ref="${xml(range)}"/>`).join('')}</mergeCells>` : ''}<pageMargins${attributes(setup.margins ?? { left: 0.7, right: 0.7, top: 0.75, bottom: 0.75, header: 0.3, footer: 0.3 })}/><pageSetup${attributes({ orientation: setup.orientation, paperSize: setup.paperSize, fitToWidth: setup.fitToWidth, fitToHeight: setup.fitToHeight })}/>${sheet.headerFooter.oddFooter ? `<headerFooter><oddFooter>${xml(sheet.headerFooter.oddFooter)}</oddFooter></headerFooter>` : ''}</worksheet>`;
}

function relationships(
	entries: Array<{ Id: string; Type: string; Target: string }>
): string {
	return `${declaration}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${entries.map((entry) => `<Relationship${attributes(entry)}/>`).join('')}</Relationships>`;
}

function workbookArchive(book: Workbook): Map<string, string> {
	if (!book.worksheets.length)
		throw new XlsxError('A workbook needs at least one worksheet.');
	const zip = new Map<string, string>();
	const styles = new StyleTable();
	const strings = new StringTable();
	for (const [index, sheet] of book.worksheets.entries())
		zip.set(
			`xl/worksheets/sheet${index + 1}.xml`,
			worksheetXml(sheet, styles, strings)
		);
	zip.set('xl/styles.xml', styles.toXml());
	zip.set('xl/sharedStrings.xml', strings.toXml());
	const names: string[] = [];
	for (const [index, sheet] of book.worksheets.entries()) {
		const prefix = `'${sheet.name.replace(/'/g, "''")}'!`;
		if (sheet.pageSetup.printArea)
			names.push(
				`<definedName name="_xlnm.Print_Area" localSheetId="${index}">${xml(prefix + sheet.pageSetup.printArea.replace(/([A-Z]+)(\d+)/g, '$$$1$$$2'))}</definedName>`
			);
		if (sheet.pageSetup.printTitlesRow)
			names.push(
				`<definedName name="_xlnm.Print_Titles" localSheetId="${index}">${xml(
					prefix +
						sheet.pageSetup.printTitlesRow
							.split(':')
							.map((row) => '$' + row)
							.join(':')
				)}</definedName>`
			);
	}
	zip.set(
		'xl/workbook.xml',
		`${declaration}<workbook xmlns="${spreadsheetNamespace}" xmlns:r="${relationshipNamespace}"><bookViews><workbookView/></bookViews><sheets>${book.worksheets.map((sheet, index) => `<sheet name="${xml(sheet.name)}" sheetId="${index + 1}" r:id="rId${index + 1}"/>`).join('')}</sheets>${names.length ? `<definedNames>${names.join('')}</definedNames>` : ''}<calcPr calcId="171027" fullCalcOnLoad="${Number(book.calcProperties.fullCalcOnLoad)}"/></workbook>`
	);
	zip.set(
		'xl/_rels/workbook.xml.rels',
		relationships([
			...book.worksheets.map((_, index) => ({
				Id: `rId${index + 1}`,
				Type: `${relationshipNamespace}/worksheet`,
				Target: `worksheets/sheet${index + 1}.xml`
			})),
			{
				Id: 'styles',
				Type: `${relationshipNamespace}/styles`,
				Target: 'styles.xml'
			},
			{
				Id: 'strings',
				Type: `${relationshipNamespace}/sharedStrings`,
				Target: 'sharedStrings.xml'
			}
		])
	);
	zip.set(
		'_rels/.rels',
		relationships([
			{
				Id: 'workbook',
				Type: `${relationshipNamespace}/officeDocument`,
				Target: 'xl/workbook.xml'
			},
			{
				Id: 'core',
				Type: 'http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties',
				Target: 'docProps/core.xml'
			},
			{
				Id: 'app',
				Type: `${relationshipNamespace}/extended-properties`,
				Target: 'docProps/app.xml'
			}
		])
	);
	zip.set(
		'docProps/core.xml',
		`${declaration}<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"><dc:creator>${xml(book.creator)}</dc:creator><cp:lastModifiedBy>${xml(book.creator)}</cp:lastModifiedBy><dcterms:created xsi:type="dcterms:W3CDTF">${book.created.toISOString()}</dcterms:created><dcterms:modified xsi:type="dcterms:W3CDTF">${book.modified.toISOString()}</dcterms:modified></cp:coreProperties>`
	);
	zip.set(
		'docProps/app.xml',
		`${declaration}<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"><Application>MinkExcel</Application></Properties>`
	);
	const overrides = [
		{
			PartName: '/xl/workbook.xml',
			ContentType:
				'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml'
		},
		{
			PartName: '/xl/sharedStrings.xml',
			ContentType:
				'application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml'
		},
		{
			PartName: '/xl/styles.xml',
			ContentType:
				'application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml'
		},
		{
			PartName: '/docProps/core.xml',
			ContentType: 'application/vnd.openxmlformats-package.core-properties+xml'
		},
		{
			PartName: '/docProps/app.xml',
			ContentType:
				'application/vnd.openxmlformats-officedocument.extended-properties+xml'
		},
		...book.worksheets.map((_, index) => ({
			PartName: `/xl/worksheets/sheet${index + 1}.xml`,
			ContentType:
				'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml'
		}))
	];
	zip.set(
		'[Content_Types].xml',
		`${declaration}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>${overrides.map((entry) => `<Override${attributes(entry)}/>`).join('')}</Types>`
	);
	return zip;
}

export async function writeWorkbook(
	book: Workbook,
	signal?: AbortSignal
): Promise<Uint8Array> {
	await checkpoint(signal);
	try {
		return await writeZip(workbookArchive(book), signal);
	} catch (cause) {
		if (signal?.aborted) throw signal.reason;
		throw cause instanceof XlsxError
			? cause
			: new XlsxError('Cannot write workbook.', { cause });
	}
}
