import { ZipArchive, type ArchiveLimits } from './zip.js';
import { checkpoint } from './async.js';
import { Workbook, XlsxError, XlsxLimitError, parseAddress, type CellValue, type Worksheet } from './model.js';
import { child, children, decodeExcelText, parseXml, stringText, type XmlNode, type XmlVisitor } from './xml.js';
import { templates, valueKey, metadataKey, presentationKey, type TemplateSheet, type TemplateRow } from './template.js';
type Limits = ArchiveLimits & { rows: number; columns: number; cells: number };
export type ReadLimits = Partial<Limits>;
export type ReadOptions = ReadLimits & { preserveTemplate?: boolean };
const defaultLimits: Limits = {
	fileBytes: 5 * 1024 * 1024,
	entries: 1_000,
	entryBytes: 20 * 1024 * 1024,
	totalBytes: 50 * 1024 * 1024,
	compressionRatio: 200,
	rows: 10_000,
	columns: 100,
	cells: 100_000
};
type Relationship = { id: string; type: string; target: string };

/** Resolve only package-local targets; no filesystem reads or network requests. */
function targetPath(owner: string, target: string): string {
	if (!target || /[\\:#?]/.test(target))
		throw new XlsxError('Invalid workbook relationship target.');
	const parts = (
		target.startsWith('/') ? target.slice(1) : owner.slice(0, owner.lastIndexOf('/') + 1) + target
	).split('/');
	const resolved: string[] = [];
	for (const part of parts) {
		if (part === '..') {
			if (!resolved.length) throw new XlsxError('Workbook relationship leaves the archive.');
			resolved.pop();
		} else if (part && part !== '.') resolved.push(part);
	}
	return resolved.join('/');
}

function relsPath(owner: string): string {
	const slash = owner.lastIndexOf('/');
	return owner.slice(0, slash + 1) + '_rels/' + owner.slice(slash + 1) + '.rels';
}
function relationships(root: XmlNode, owner: string): Relationship[] {
	if (root.name !== 'Relationships') throw new XlsxError('Invalid workbook relationships.');
	return children(root, 'Relationship')
		.filter((entry) => entry.attributes.TargetMode !== 'External')
		.map((entry) => ({
			id: entry.attributes.Id,
			type: entry.attributes.Type,
			target: targetPath(owner, entry.attributes.Target)
		}));
}
function dateFormat(format: string): boolean {
	return /[ymdhs]/i.test(format.replace(/"[^"]*"|\\.|\[(?![hms]+\])[^\]]*\]/gi, ''));
}
function serialDate(serial: number, date1904: boolean): Date {
	const time =
		(date1904 ? Date.UTC(1904, 0, 1) : Date.UTC(1899, 11, 30)) +
		(serial + (!date1904 && serial < 60 ? 1 : 0)) * 86_400_000;
	const date = new Date(Math.round(time));
	if (!Number.isFinite(date.getTime())) throw new XlsxError('Invalid date cell.');
	return date;
}

type CellContext = { strings: string[]; dateStyles: boolean[]; date1904: boolean };

function isoDate(text: string): Date {
	const date = new Date(text);
	if (!Number.isFinite(date.getTime())) throw new XlsxError('Invalid date cell.');
	return date;
}
function sharedString(raw: string | undefined, strings: string[]): string {
	const index = Number(raw);
	if (raw === undefined || !Number.isInteger(index) || index < 0 || index >= strings.length)
		throw new XlsxError('Invalid shared-string index.');
	return strings[index];
}
function numberValue(raw: string | undefined): number | null {
	if (raw === undefined || raw === '') return null;
	const value = Number(raw);
	if (!Number.isFinite(value)) throw new XlsxError('Invalid numeric cell.');
	return value;
}
function scalarValue(node: XmlNode, strings: string[]): CellValue {
	const raw = child(node, 'v')?.text;
	switch (node.attributes.t) {
		case 'inlineStr': {
			const inline = child(node, 'is');
			return inline ? stringText(inline) : '';
		}
		case 's': return sharedString(raw, strings);
		case 'str': return decodeExcelText(raw ?? '');
		case 'b':
			if (raw !== '0' && raw !== '1') throw new XlsxError('Invalid boolean cell.');
			return raw === '1';
		case 'e': return { error: raw ?? '#VALUE!' };
		case 'd': return isoDate(raw ?? '');
		case 'n':
		case undefined: return numberValue(raw);
		default: throw new XlsxError('Unsupported workbook cell type.');
	}
}
function cellValue(node: XmlNode, context: CellContext): CellValue {
	const value = scalarValue(node, context.strings);
	const formula = child(node, 'f');
	if (formula) {
		if (value === null || value instanceof Date) return { formula: formula.text };
		return { formula: formula.text, result: value as string | number | boolean | { error: string } };
	}
	if (typeof value === 'number' && context.dateStyles[Number(node.attributes.s ?? 0)])
		return serialDate(value, context.date1904);
	return value;
}

function importLimits(bytes: Uint8Array, options: ReadOptions): Limits {
	if (options.preserveTemplate !== undefined && typeof options.preserveTemplate !== 'boolean')
		throw new XlsxError('preserveTemplate must be a boolean.');
	const limits = { ...defaultLimits };
	for (const key of Object.keys(defaultLimits) as Array<keyof Limits>) {
		const value = options[key];
		if (value === undefined) continue;
		const valid = key === 'compressionRatio'
			? Number.isFinite(value) && value > 0
			: Number.isSafeInteger(value) && value >= 0;
		if (!valid) throw new XlsxLimitError(`Invalid import limit: ${key}.`);
		limits[key] = value;
	}
	if (bytes.byteLength > limits.fileBytes)
		throw new XlsxLimitError(`File exceeds the ${limits.fileBytes} bytes import limit.`);
	return limits;
}

class ImportContext {
	readonly zip: ZipArchive;
	readonly parts = new Map<string, Uint8Array>();
	private sources = new Map<string, string>();
	private remaining: number;
	constructor(bytes: Uint8Array, readonly limits: Limits, readonly preserve: boolean, readonly signal?: AbortSignal) {
		this.zip = new ZipArchive(bytes, limits);
		this.remaining = limits.totalBytes;
	}
	async readBytes(path: string): Promise<Uint8Array> {
		const cached = this.parts.get(path);
		if (cached) return cached;
		const data = await this.zip.read(path, Math.min(this.limits.entryBytes, this.remaining), this.signal);
		this.remaining -= data.length;
		if (this.preserve) this.parts.set(path, data.slice());
		return data;
	}
	async readPart(path: string, onClose?: XmlVisitor): Promise<XmlNode> {
		const source = new TextDecoder('utf-8', { fatal: true }).decode(await this.readBytes(path));
		if (this.preserve) this.sources.set(path, source.replace(/^\uFEFF/, '').replace(/\r\n?/g, '\n'));
		return parseXml(source, onClose, this.preserve);
	}
	retainedPart(path: string, document: XmlNode) {
		const source = this.sources.get(path);
		if (source === undefined) throw new XlsxError('Missing retained template XML.');
		return { path, xml: source, document };
	}
}

async function sharedStrings(rels: Relationship[], context: ImportContext): Promise<string[]> {
	const relation = rels.find((entry) => entry.type.endsWith('/sharedStrings'));
	if (!relation) return [];
	const strings: string[] = [];
	const document = await context.readPart(relation.target, (node, parent, depth) => {
		if (node.name !== 'si' || parent?.name !== 'sst' || depth !== 1) return false;
		strings.push(stringText(node));
		return true;
	});
	if (document.name !== 'sst') throw new XlsxError('Invalid shared-string document.');
	return strings;
}
async function dateFormats(rels: Relationship[], context: ImportContext): Promise<boolean[]> {
	const relation = rels.find((entry) => entry.type.endsWith('/styles'));
	if (!relation) return [];
	const styles = await context.readPart(relation.target);
	const customFormats = new Map((child(styles, 'numFmts')?.children ?? [])
		.map((entry) => [Number(entry.attributes.numFmtId), entry.attributes.formatCode]));
	return (child(styles, 'cellXfs')?.children ?? []).map((entry) => {
		const id = Number(entry.attributes.numFmtId);
		return (id >= 14 && id <= 22) || (id >= 45 && id <= 47) || dateFormat(customFormats.get(id) ?? '');
	});
}
function worksheetRelation(node: XmlNode, rels: Relationship[]): Relationship {
	// Namespace prefixes can differ between producers; match ids by local name.
	const id = Object.entries(node.attributes).find(
		([name]) => name.includes(':') && name.split(':').at(-1) === 'id')?.[1];
	const relation = rels.find((entry) => entry.id === id && entry.type.endsWith('/worksheet'));
	if (!relation) throw new XlsxError('Missing worksheet relationship.');
	return relation;
}

class WorksheetReader {
	readonly rows = new Map<number, TemplateRow>();
	readonly values = new Map<string, string>();
	private previousRow = 0;
	private maximumColumn = 0;
	constructor(readonly sheet: Worksheet, private context: ImportContext, private cells: CellContext) {}
	private checkExtent(row: number): void {
		const limit = this.context.limits.cells;
		if (row * Math.max(this.maximumColumn, 1) > limit)
			throw new XlsxLimitError(`The worksheet exceeds the ${limit}-cell limit.`);
	}
	private cellAddress(node: XmlNode, row: number, previousColumn: number) {
		const address = node.attributes.r ? parseAddress(node.attributes.r) : { row, column: previousColumn + 1 };
		if (address.row !== row || address.column <= previousColumn)
			throw new XlsxError('Invalid worksheet cell order.');
		if (address.column > this.context.limits.columns)
			throw new XlsxLimitError(`The worksheet exceeds the ${this.context.limits.columns}-column limit.`);
		this.maximumColumn = Math.max(this.maximumColumn, address.column);
		this.checkExtent(row);
		return address;
	}
	readRow(node: XmlNode): void {
		const number = node.attributes.r === undefined ? this.previousRow + 1 : Number(node.attributes.r);
		if (number <= this.previousRow) throw new XlsxError('Invalid worksheet row order.');
		if (number - 1 > this.context.limits.rows)
			throw new XlsxLimitError(`The worksheet exceeds the ${this.context.limits.rows}-row limit.`);
		this.checkExtent(number);
		this.previousRow = number;
		const row = this.sheet.getRow(number);
		const retained = this.context.preserve ? new Map<number, XmlNode>() : undefined;
		let previousColumn = 0;
		for (const cellNode of children(node, 'c')) {
			const address = this.cellAddress(cellNode, number, previousColumn);
			previousColumn = address.column;
			const cell = row.getCell(address.column);
			cell.value = cellValue(cellNode, this.cells);
			if (retained) {
				retained.set(address.column, cellNode);
				this.values.set(cell.address, valueKey(cell.value));
			}
		}
		if (retained) this.rows.set(number, { node, cells: retained });
	}
	async read(path: string): Promise<XmlNode> {
		const retained = this.context.signal || this.context.preserve;
		// Value-only imports consume completed rows without retaining the XML tree.
		const visit: XmlVisitor = (node, parent, depth) => {
			if (node.name !== 'row' || parent?.name !== 'sheetData' || depth !== 2) return false;
			this.readRow(node);
			return true;
		};
		const document = await this.context.readPart(path, retained ? undefined : visit);
		if (document.name !== 'worksheet') throw new XlsxError('Invalid worksheet document.');
		const data = children(document, 'sheetData');
		if (data.length !== 1) throw new XlsxError('Invalid or missing worksheet data.');
		if (retained) await this.readRows(data[0]);
		return document;
	}
	private async readRows(data: XmlNode): Promise<void> {
		for (const node of children(data, 'row')) {
			if (this.previousRow % 256 === 0) await checkpoint(this.context.signal);
			this.readRow(node);
		}
	}
}

function coreMetadata(book: Workbook, document: XmlNode): void {
	book.creator = child(document, 'creator')?.text ?? book.creator;
	book.lastModifiedBy = child(document, 'lastModifiedBy')?.text;
	for (const name of ['created', 'modified'] as const) {
		const text = child(document, name)?.text;
		if (!text) continue;
		const date = new Date(text);
		if (!Number.isFinite(date.getTime())) throw new XlsxError('Invalid template metadata date.');
		book[name] = date;
	}
}
async function retainTemplate(book: Workbook, sheets: TemplateSheet[], cells: CellContext,
	workbook: ReturnType<ImportContext['retainedPart']>, rootRels: Relationship[], context: ImportContext): Promise<void> {
	const relation = rootRels.find((entry) => entry.type.endsWith('/core-properties'));
	let core: ReturnType<ImportContext['retainedPart']> | undefined;
	if (relation) {
		const document = await context.readPart(relation.target);
		core = context.retainedPart(relation.target, document);
		coreMetadata(book, document);
	}
	// Validate CRCs for every retained part, including binary media and comments.
	for (const name of context.zip.names()) await context.readBytes(name);
	templates.set(book, {
		parts: context.parts, sheets, date1904: cells.date1904, dateStyles: cells.dateStyles,
		workbook, core, metadata: metadataKey(book), calculation: JSON.stringify(book.calcProperties)
	});
}
async function importWorkbook(context: ImportContext): Promise<Workbook> {
	const rootRels = relationships(await context.readPart('_rels/.rels'), '');
	const workbookRel = rootRels.find((entry) => entry.type.endsWith('/officeDocument'));
	if (!workbookRel) throw new XlsxError('Missing workbook relationship.');
	const document = await context.readPart(workbookRel.target);
	if (document.name !== 'workbook') throw new XlsxError('Invalid workbook document.');
	const rels = relationships(await context.readPart(relsPath(workbookRel.target)), workbookRel.target);
	const cells: CellContext = {
		strings: await sharedStrings(rels, context),
		dateStyles: await dateFormats(rels, context),
		date1904: ['1', 'true'].includes(child(document, 'workbookPr')?.attributes.date1904 ?? '')
	};
	const book = new Workbook();
	const retainedSheets: TemplateSheet[] = [];
	const sheets = child(document, 'sheets');
	if (!sheets) throw new XlsxError('Missing workbook sheets.');
	for (const node of children(sheets, 'sheet')) {
		const relation = worksheetRelation(node, rels);
		const reader = new WorksheetReader(book.addWorksheet(node.attributes.name), context, cells);
		const document = await reader.read(relation.target);
		if (context.preserve) retainedSheets.push({
			...context.retainedPart(relation.target, document), sheet: reader.sheet,
			rows: reader.rows, values: reader.values, presentation: presentationKey(reader.sheet)
		});
	}
	if (context.preserve)
		await retainTemplate(book, retainedSheets, cells, context.retainedPart(workbookRel.target, document), rootRels, context);
	return book;
}

/** Value-oriented import; preserveTemplate retains the archive for cell value edits. */
export async function readWorkbook(bytes: Uint8Array, options: ReadOptions = {}, signal?: AbortSignal): Promise<Workbook> {
	signal?.throwIfAborted();
	try {
		const context = new ImportContext(bytes, importLimits(bytes, options), options.preserveTemplate === true, signal);
		return await importWorkbook(context);
	} catch (cause) {
		if (signal?.aborted) throw signal.reason;
		throw cause instanceof XlsxError ? cause : new XlsxError('Invalid workbook.', { cause });
	}
}
