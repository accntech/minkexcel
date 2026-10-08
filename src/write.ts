import { writeZip } from './zip.js';
import { checkpoint } from './async.js';
import { templateArchive } from './template.js';
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

function fontXml(font: NonNullable<Style['font']>): string {
	return `<font>${font.bold ? '<b/>' : ''}${font.italic ? '<i/>' : ''}${font.size !== undefined ? `<sz val="${xml(font.size)}"/>` : ''}${font.name ? `<name val="${xml(font.name)}"/>` : ''}${font.color ? `<color rgb="${xml(font.color.argb)}"/>` : ''}</font>`;
}
function borderXml(borders: NonNullable<Style['border']>): string {
	const edges = (['left', 'right', 'top', 'bottom'] as const).map((side) => {
		const border = borders[side];
		if (!border) return `<${side}/>`;
		const color = border.color ? `<color rgb="${xml(border.color.argb)}"/>` : '';
		return `<${side}${attributes({ style: border.style })}>${color}</${side}>`;
	});
	return `<border>${edges.join('')}<diagonal/></border>`;
}
function alignmentXml(alignment: Style['alignment']): string {
	if (!alignment || !Object.keys(alignment).length) return '';
	return `<alignment${attributes({ ...alignment, vertical: alignment.vertical === 'middle' ? 'center' : alignment.vertical })}/>`;
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
	private borders: string[] = ['<border><left/><right/><top/><bottom/><diagonal/></border>'];
	private records: string[] = [
		'<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>'
	];
	private register(entries: string[], entry: string): number {
		const index = entries.indexOf(entry);
		if (index >= 0) return index;
		entries.push(entry);
		return entries.length - 1;
	}
	private fontId(font: Style['font']): number {
		if (!font || !Object.keys(font).length) return 0;
		return this.register(this.fonts, fontXml(font));
	}
	private fillId(fill: Style['fill']): number {
		if (!fill) return 0;
		return this.register(this.fills,
			`<fill><patternFill patternType="solid"><fgColor rgb="${xml(fill.fgColor.argb)}"/><bgColor indexed="64"/></patternFill></fill>`);
	}
	private formatId(format: Style['numFmt']): number {
		if (!format || format === 'General') return 0;
		return 164 + this.register(this.formats, format);
	}
	private borderId(border: Style['border']): number {
		if (!border || !Object.keys(border).length) return 0;
		return this.register(this.borders, borderXml(border));
	}
	id(style: Style): number {
		const fontId = this.fontId(style.font);
		const fillId = this.fillId(style.fill);
		const numFmtId = this.formatId(style.numFmt);
		const borderId = this.borderId(style.border);
		const alignment = alignmentXml(style.alignment);
		if (!fontId && !fillId && !numFmtId && !borderId && !alignment) return 0;
		const attrs = attributes({
			numFmtId, fontId, fillId, borderId, xfId: 0,
			applyFont: fontId ? true : undefined,
			applyFill: fillId ? true : undefined,
			applyBorder: borderId ? true : undefined,
			applyNumberFormat: numFmtId ? true : undefined,
			applyAlignment: alignment ? true : undefined
		});
		return this.register(this.records, `<xf${attrs}>${alignment}</xf>`);
	}
	toXml(): string {
		return `${declaration}<styleSheet xmlns="${spreadsheetNamespace}"><numFmts count="${this.formats.length}">${this.formats.map((format, index) => `<numFmt numFmtId="${index + 164}" formatCode="${xml(format)}"/>`).join('')}</numFmts><fonts count="${this.fonts.length}">${this.fonts.join('')}</fonts><fills count="${this.fills.length}">${this.fills.join('')}</fills><borders count="${this.borders.length}">${this.borders.join('')}</borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="${this.records.length}">${this.records.join('')}</cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>`;
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

function formulaCache(result: { formula: string; result?: string | number | boolean | { error: string } }['result']): { type?: string; content: string } {
	if (result === undefined) return { content: '' };
	if (typeof result === 'number') return { content: `<v>${numeric(result)}</v>` };
	if (typeof result === 'boolean') return { type: 'b', content: `<v>${Number(result)}</v>` };
	if (typeof result === 'string') return { type: 'str', content: `<v>${excelText(result)}</v>` };
	return { type: 'e', content: `<v>${xml(result.error)}</v>` };
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
	const cached = formulaCache(value.result);
	return `<c${attrs}${attributes({ t: cached.type })}><f>${xml(value.formula)}</f>${cached.content}</c>`;
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
	return `${declaration}<worksheet xmlns="${spreadsheetNamespace}" xmlns:r="${relationshipNamespace}">${setup.fitToPage ? '<sheetPr><pageSetUpPr fitToPage="1"/></sheetPr>' : ''}<dimension ref="A1:${cellAddress(Math.max(sheet.rowCount, 1), Math.max(sheet.columnCount, 1))}"/>${views ? `<sheetViews>${views}</sheetViews>` : ''}<sheetFormatPr defaultRowHeight="15"/>${columnXml ? `<cols>${columnXml}</cols>` : ''}<sheetData>${data}</sheetData>${sheet.autoFilter ? `<autoFilter ref="${xml(address(sheet.autoFilter.from) + ':' + address(sheet.autoFilter.to))}"/>` : ''}${sheet.merges.length ? `<mergeCells count="${sheet.merges.length}">${sheet.merges.map((range) => `<mergeCell ref="${xml(range)}"/>`).join('')}</mergeCells>` : ''}${setup.horizontalCentered !== undefined ? `<printOptions${attributes({ horizontalCentered: setup.horizontalCentered })}/>` : ''}<pageMargins${attributes(setup.margins ?? { left: 0.7, right: 0.7, top: 0.75, bottom: 0.75, header: 0.3, footer: 0.3 })}/><pageSetup${attributes({ orientation: setup.orientation, paperSize: setup.paperSize, fitToWidth: setup.fitToWidth, fitToHeight: setup.fitToHeight })}/>${sheet.headerFooter.oddFooter ? `<headerFooter><oddFooter>${xml(sheet.headerFooter.oddFooter)}</oddFooter></headerFooter>` : ''}</worksheet>`;
}

function relationships(
	entries: Array<{ Id: string; Type: string; Target: string }>
): string {
	return `${declaration}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${entries.map((entry) => `<Relationship${attributes(entry)}/>`).join('')}</Relationships>`;
}

function definedNames(book: Workbook): string {
	const names: string[] = [];
	for (const [index, sheet] of book.worksheets.entries()) {
		const prefix = `'${sheet.name.replace(/'/g, "''")}'!`;
		if (sheet.pageSetup.printArea) {
			const area = sheet.pageSetup.printArea.replace(/([A-Z]+)(\d+)/g, '$$$1$$$2');
			names.push(`<definedName name="_xlnm.Print_Area" localSheetId="${index}">${xml(prefix + area)}</definedName>`);
		}
		if (sheet.pageSetup.printTitlesRow) {
			const rows = sheet.pageSetup.printTitlesRow.split(':').map((row) => '$' + row).join(':');
			names.push(`<definedName name="_xlnm.Print_Titles" localSheetId="${index}">${xml(prefix + rows)}</definedName>`);
		}
	}
	return names.length ? `<definedNames>${names.join('')}</definedNames>` : '';
}
function workbookDocument(book: Workbook): string {
	const sheets = book.worksheets.map((sheet, index) =>
		`<sheet name="${xml(sheet.name)}" sheetId="${index + 1}" r:id="rId${index + 1}"/>`).join('');
	return `${declaration}<workbook xmlns="${spreadsheetNamespace}" xmlns:r="${relationshipNamespace}"><bookViews><workbookView/></bookViews><sheets>${sheets}</sheets>${definedNames(book)}<calcPr calcId="171027" fullCalcOnLoad="${Number(book.calcProperties.fullCalcOnLoad)}"/></workbook>`;
}
function workbookRelationships(book: Workbook): string {
	return relationships([
		...book.worksheets.map((_, index) => ({
			Id: `rId${index + 1}`, Type: `${relationshipNamespace}/worksheet`, Target: `worksheets/sheet${index + 1}.xml`
		})),
		{ Id: 'styles', Type: `${relationshipNamespace}/styles`, Target: 'styles.xml' },
		{ Id: 'strings', Type: `${relationshipNamespace}/sharedStrings`, Target: 'sharedStrings.xml' }
	]);
}
function rootRelationships(): string {
	return relationships([
		{ Id: 'workbook', Type: `${relationshipNamespace}/officeDocument`, Target: 'xl/workbook.xml' },
		{
			Id: 'core', Type: 'http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties',
			Target: 'docProps/core.xml'
		},
		{ Id: 'app', Type: `${relationshipNamespace}/extended-properties`, Target: 'docProps/app.xml' }
	]);
}
function coreProperties(book: Workbook): string {
	return `${declaration}<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"><dc:creator>${xml(book.creator)}</dc:creator><cp:lastModifiedBy>${xml(book.lastModifiedBy ?? book.creator)}</cp:lastModifiedBy><dcterms:created xsi:type="dcterms:W3CDTF">${book.created.toISOString()}</dcterms:created><dcterms:modified xsi:type="dcterms:W3CDTF">${book.modified.toISOString()}</dcterms:modified></cp:coreProperties>`;
}
function contentTypes(book: Workbook): string {
	const overrides = [
		{ PartName: '/xl/workbook.xml', ContentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml' },
		{ PartName: '/xl/sharedStrings.xml', ContentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml' },
		{ PartName: '/xl/styles.xml', ContentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml' },
		{ PartName: '/docProps/core.xml', ContentType: 'application/vnd.openxmlformats-package.core-properties+xml' },
		{ PartName: '/docProps/app.xml', ContentType: 'application/vnd.openxmlformats-officedocument.extended-properties+xml' },
		...book.worksheets.map((_, index) => ({
			PartName: `/xl/worksheets/sheet${index + 1}.xml`,
			ContentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml'
		}))
	];
	return `${declaration}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>${overrides.map((entry) => `<Override${attributes(entry)}/>`).join('')}</Types>`;
}
function workbookArchive(book: Workbook): Map<string, string> {
	if (!book.worksheets.length) throw new XlsxError('A workbook needs at least one worksheet.');
	const zip = new Map<string, string>();
	const styles = new StyleTable();
	const strings = new StringTable();
	for (const [index, sheet] of book.worksheets.entries())
		zip.set(`xl/worksheets/sheet${index + 1}.xml`, worksheetXml(sheet, styles, strings));
	zip.set('xl/styles.xml', styles.toXml());
	zip.set('xl/sharedStrings.xml', strings.toXml());
	zip.set('xl/workbook.xml', workbookDocument(book));
	zip.set('xl/_rels/workbook.xml.rels', workbookRelationships(book));
	zip.set('_rels/.rels', rootRelationships());
	zip.set('docProps/core.xml', coreProperties(book));
	zip.set('docProps/app.xml', `${declaration}<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"><Application>MinkExcel</Application></Properties>`);
	zip.set('[Content_Types].xml', contentTypes(book));
	return zip;
}

export async function writeWorkbook(
	book: Workbook,
	signal?: AbortSignal
): Promise<Uint8Array> {
	await checkpoint(signal);
	try {
		return await writeZip(templateArchive(book) ?? workbookArchive(book), signal);
	} catch (cause) {
		if (signal?.aborted) throw signal.reason;
		throw cause instanceof XlsxError
			? cause
			: new XlsxError('Cannot write workbook.', { cause });
	}
}
