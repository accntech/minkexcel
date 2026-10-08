import { ZipArchive, type ArchiveLimits } from './zip.js';
import { checkpoint } from './async.js';
import { Workbook, XlsxError, XlsxLimitError, parseAddress, type CellValue } from './model.js';
import { child, children, decodeExcelText, parseXml, stringText, type XmlNode, type XmlVisitor } from './xml.js';
type Limits = ArchiveLimits & { rows: number; columns: number; cells: number };
export type ReadLimits = Partial<Limits>;
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

function cellValue(
	node: XmlNode,
	strings: string[],
	dateStyles: boolean[],
	date1904: boolean
): CellValue {
	const raw = child(node, 'v')?.text;
	const formula = child(node, 'f');
	let value: CellValue;
	switch (node.attributes.t) {
		case 'inlineStr': {
			const inline = child(node, 'is');
			value = inline ? stringText(inline) : '';
			break;
		}
		case 's': {
			const index = Number(raw);
			if (raw === undefined || !Number.isInteger(index) || index < 0 || index >= strings.length)
				throw new XlsxError('Invalid shared-string index.');
			value = strings[index];
			break;
		}
		case 'str':
			value = decodeExcelText(raw ?? '');
			break;
		case 'b':
			if (raw !== '0' && raw !== '1') throw new XlsxError('Invalid boolean cell.');
			value = raw === '1';
			break;
		case 'e':
			value = { error: raw ?? '#VALUE!' };
			break;
		case 'd': {
			const date = new Date(raw ?? '');
			if (!Number.isFinite(date.getTime())) throw new XlsxError('Invalid date cell.');
			value = date;
			break;
		}
		case 'n':
		case undefined: {
			if (raw === undefined || raw === '') value = null;
			else {
				const number = Number(raw);
				if (!Number.isFinite(number)) throw new XlsxError('Invalid numeric cell.');
				value =
					!formula && dateStyles[Number(node.attributes.s ?? 0)]
						? serialDate(number, date1904)
						: number;
			}
			break;
		}
		default:
			throw new XlsxError('Unsupported workbook cell type.');
	}
	if (formula)
		return {
			formula: formula.text,
			...(value !== null && !(value instanceof Date)
				? { result: value as string | number | boolean | { error: string } }
				: {})
		};
	return value;
}

/** Value-oriented import reader. Imported presentation features are deliberately not round-tripped. */
export async function readWorkbook(
	bytes: Uint8Array,
	options: ReadLimits = {},
	signal?: AbortSignal
): Promise<Workbook> {
	signal?.throwIfAborted();
	try {
		const limits = { ...defaultLimits };
		for (const key of Object.keys(defaultLimits) as Array<keyof Limits>) {
			const value = options[key];
			if (value === undefined) continue;
			if (key === 'compressionRatio'
				? !Number.isFinite(value) || value <= 0
				: !Number.isSafeInteger(value) || value < 0)
				throw new XlsxLimitError(`Invalid import limit: ${key}.`);
			limits[key] = value;
		}
		if (bytes.byteLength > limits.fileBytes)
			throw new XlsxLimitError(`File exceeds the ${limits.fileBytes} bytes import limit.`);
		const zip = new ZipArchive(bytes, limits);
		let remaining = limits.totalBytes;
		const readPart = async (path: string, onClose?: XmlVisitor) => {
			const data = await zip.read(path, Math.min(limits.entryBytes, remaining), signal);
			remaining -= data.length;
			return parseXml(new TextDecoder('utf-8', { fatal: true }).decode(data), onClose);
		};
		const rootRels = relationships(await readPart('_rels/.rels'), '');
		const workbookRel = rootRels.find((entry) => entry.type.endsWith('/officeDocument'));
		if (!workbookRel) throw new XlsxError('Missing workbook relationship.');
		const document = await readPart(workbookRel.target);
		if (document.name !== 'workbook') throw new XlsxError('Invalid workbook document.');
		const rels = relationships(await readPart(relsPath(workbookRel.target)), workbookRel.target);
		const shared = rels.find((entry) => entry.type.endsWith('/sharedStrings'));
		const strings: string[] = [];
		if (shared) {
			const document = await readPart(shared.target, (node, parent, depth) => {
				if (node.name !== 'si' || parent?.name !== 'sst' || depth !== 1) return false;
				strings.push(stringText(node));
				return true;
			});
			if (document.name !== 'sst') throw new XlsxError('Invalid shared-string document.');
		}
		const styleRel = rels.find((entry) => entry.type.endsWith('/styles'));
		let dateStyles: boolean[] = [];
		if (styleRel) {
			const styles = await readPart(styleRel.target);
			const customFormats = new Map(
				(child(styles, 'numFmts')?.children ?? []).map((entry) => [
					Number(entry.attributes.numFmtId),
					entry.attributes.formatCode
				])
			);
			dateStyles = (child(styles, 'cellXfs')?.children ?? []).map((entry) => {
				const id = Number(entry.attributes.numFmtId);
				return (
					(id >= 14 && id <= 22) ||
					(id >= 45 && id <= 47) ||
					dateFormat(customFormats.get(id) ?? '')
				);
			});
		}
		const date1904 = ['1', 'true'].includes(
			child(document, 'workbookPr')?.attributes.date1904 ?? ''
		);
		const book = new Workbook();
		const sheets = child(document, 'sheets');
		if (!sheets) throw new XlsxError('Missing workbook sheets.');
		for (const sheetNode of children(sheets, 'sheet')) {
			// Namespace prefixes can differ between producers; match the relationship id by local name.
			const id = Object.entries(sheetNode.attributes).find(
				([name]) => name.includes(':') && name.split(':').at(-1) === 'id'
			)?.[1];
			const relation = rels.find((entry) => entry.id === id && entry.type.endsWith('/worksheet'));
			if (!relation) throw new XlsxError('Missing worksheet relationship.');
			const sheet = book.addWorksheet(sheetNode.attributes.name);
			let previousRow = 0;
			let maximumColumn = 0;
			const readRow = (rowNode: XmlNode) => {
				const rowNumber =
					rowNode.attributes.r === undefined ? previousRow + 1 : Number(rowNode.attributes.r);
				if (rowNumber <= previousRow) throw new XlsxError('Invalid worksheet row order.');
				if (rowNumber - 1 > limits.rows)
					throw new XlsxLimitError(`The worksheet exceeds the ${limits.rows}-row limit.`);
				if (rowNumber * Math.max(maximumColumn, 1) > limits.cells)
					throw new XlsxLimitError(`The worksheet exceeds the ${limits.cells}-cell limit.`);
				previousRow = rowNumber;
				const row = sheet.getRow(rowNumber);
				let previousColumn = 0;
				for (const node of children(rowNode, 'c')) {
					const address = node.attributes.r
						? parseAddress(node.attributes.r)
						: { row: rowNumber, column: previousColumn + 1 };
					if (address.row !== rowNumber || address.column <= previousColumn)
						throw new XlsxError('Invalid worksheet cell order.');
					previousColumn = address.column;
					if (address.column > limits.columns)
						throw new XlsxLimitError(`The worksheet exceeds the ${limits.columns}-column limit.`);
					maximumColumn = Math.max(maximumColumn, address.column);
					if (rowNumber * maximumColumn > limits.cells)
						throw new XlsxLimitError(`The worksheet exceeds the ${limits.cells}-cell limit.`);
					row.getCell(address.column).value = cellValue(node, strings, dateStyles, date1904);
				}
			};
			// Consume completed rows before retaining a whole worksheet XML tree.
			// The signal path retains row checkpoints for event-loop cancellation.
			const document = await readPart(relation.target, signal ? undefined : (node, parent, depth) => {
				if (node.name !== 'row' || parent?.name !== 'sheetData' || depth !== 2) return false;
				readRow(node);
				return true;
			});
			if (document.name !== 'worksheet') throw new XlsxError('Invalid worksheet document.');
			const dataSections = children(document, 'sheetData');
			if (dataSections.length !== 1) throw new XlsxError('Invalid or missing worksheet data.');
			if (signal) {
				for (const rowNode of children(dataSections[0], 'row')) {
					if (previousRow % 256 === 0) await checkpoint(signal);
					readRow(rowNode);
				}
			}
		}
		return book;
	} catch (cause) {
		if (signal?.aborted) throw signal.reason;
		throw cause instanceof XlsxError ? cause : new XlsxError('Invalid workbook.', { cause });
	}
}
