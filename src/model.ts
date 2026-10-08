export type CellValue =
	| string
	| number
	| boolean
	| Date
	| null
	| { formula: string; result?: string | number | boolean | { error: string } }
	| { error: string };
export type Font = {
	name?: string;
	size?: number;
	bold?: boolean;
	italic?: boolean;
	color?: { argb: string };
};
export type Alignment = {
	horizontal?: 'left' | 'center' | 'right';
	vertical?: 'top' | 'middle' | 'bottom';
	wrapText?: boolean;
};
export type Fill = {
	type: 'pattern';
	pattern: 'solid';
	fgColor: { argb: string };
};
export type Style = {
	font?: Font;
	alignment?: Alignment;
	fill?: Fill;
	numFmt?: string;
};
export type PageSetup = {
	orientation?: 'portrait' | 'landscape';
	paperSize?: number;
	fitToPage?: boolean;
	fitToWidth?: number;
	fitToHeight?: number;
	printTitlesRow?: string;
	printArea?: string;
	margins?: {
		left: number;
		right: number;
		top: number;
		bottom: number;
		header: number;
		footer: number;
	};
};
export type WorksheetOptions = {
	views?: Array<{ state: 'frozen'; ySplit: number }>;
	pageSetup?: PageSetup;
};
type Address = { row: number; column: number };
const columnLetters = /* @__PURE__ */ new Map<number, string>();

function validateRow(row: number): void {
	if (!Number.isInteger(row) || row < 1 || row > 1_048_576)
		throw new XlsxError('Invalid worksheet row.');
}

function validateColumn(column: number): void {
	if (!Number.isInteger(column) || column < 1 || column > 16_384)
		throw new XlsxError('Invalid worksheet column.');
}

export class XlsxError extends Error {
	readonly _tag = 'XlsxError';
}

export class XlsxLimitError extends XlsxError {}

export function columnLetter(column: number): string {
	validateColumn(column);
	const cached = columnLetters.get(column);
	if (cached !== undefined) return cached;
	let result = '';
	for (let value = column; value > 0; value = Math.floor((value - 1) / 26))
		result = String.fromCharCode(65 + ((value - 1) % 26)) + result;
	columnLetters.set(column, result);
	return result;
}

export function cellAddress(row: number, column: number): string {
	validateRow(row);
	return columnLetter(column) + row;
}

export function parseAddress(address: string): Address {
	const match = /^([A-Z]+)([1-9]\d*)$/.exec(address);
	if (!match) throw new XlsxError('Invalid cell address.');
	let column = 0;
	for (let index = 0; index < match[1].length; index++)
		column = column * 26 + match[1].charCodeAt(index) - 64;
	const row = Number(match[2]);
	validateRow(row);
	validateColumn(column);
	return { row, column };
}

class Styled {
	style: Style = {};
	get font(): Font {
		return (this.style.font ??= {});
	}
	set font(value: Font) {
		this.style.font = value;
	}
	get alignment(): Alignment {
		return (this.style.alignment ??= {});
	}
	set alignment(value: Alignment) {
		this.style.alignment = value;
	}
	get numFmt(): string {
		return this.style.numFmt ?? 'General';
	}
	set numFmt(value: string) {
		this.style.numFmt = value;
	}
}

export class Column extends Styled {
	width?: number;
	constructor(readonly number: number) {
		super();
		columnLetter(number);
	}
	get letter() {
		return columnLetter(this.number);
	}
}

export class Cell extends Styled {
	private content: CellValue = null;
	master?: Cell;
	constructor(
		readonly row: Row,
		readonly column: number
	) {
		super();
	}
	get address() {
		return cellAddress(this.row.number, this.column);
	}
	get value(): CellValue {
		return this.master ? this.master.value : this.content;
	}
	set value(value: CellValue) {
		if (this.master) this.master.value = value;
		else this.content = value;
	}
	get text(): string {
		const value = this.value;
		if (value === null) return '';
		if (typeof value === 'object' && !(value instanceof Date))
			return 'formula' in value
				? typeof value.result === 'object'
					? value.result.error
					: String(value.result ?? '')
				: value.error;
		return String(value);
	}
	get resolvedStyle(): Style {
		return {
			...this.row.sheet.getColumn(this.column).style,
			...this.row.style,
			...this.master?.style,
			...this.style
		};
	}
	override get font(): Font {
		return (this.style.font ??= { ...this.resolvedStyle.font });
	}
	override set font(value: Font) {
		this.style.font = value;
	}
	override get alignment(): Alignment {
		return (this.style.alignment ??= { ...this.resolvedStyle.alignment });
	}
	override set alignment(value: Alignment) {
		this.style.alignment = value;
	}
	override get numFmt() {
		return this.resolvedStyle.numFmt ?? 'General';
	}
	override set numFmt(value: string) {
		this.style.numFmt = value;
	}
	set fill(value: Fill) {
		this.style.fill = value;
	}
}

export class Row extends Styled {
	readonly cells = new Map<number, Cell>();
	height?: number;
	constructor(
		readonly sheet: Worksheet,
		readonly number: number
	) {
		super();
		validateRow(number);
	}
	get cellCount(): number {
		return Math.max(0, ...this.cells.keys());
	}
	getCell(column: number): Cell {
		validateColumn(column);
		let cell = this.cells.get(column);
		if (!cell) {
			cell = new Cell(this, column);
			this.cells.set(column, cell);
		}
		return cell;
	}
	set values(values: unknown[]) {
		this.cells.clear();
		values.forEach((value, index) => {
			if (value === undefined || value === null) return;
			if (!(
				typeof value === 'string' ||
				typeof value === 'number' ||
				typeof value === 'boolean' ||
				value instanceof Date ||
				(typeof value === 'object' && ('formula' in value || 'error' in value))
			))
				throw new XlsxError('Unsupported cell value.');
			this.getCell(index + 1).value = value as CellValue;
		});
	}
	get values(): CellValue[] {
		const values: CellValue[] = [];
		for (const [column, cell] of this.cells)
			if (cell.value !== null) values[column] = cell.value;
		return values;
	}
	eachCell(callback: (cell: Cell, column: number) => void) {
		for (const [column, cell] of [...this.cells].sort(([a], [b]) => a - b))
			if (cell.value !== null) callback(cell, column);
	}
}

export class Worksheet {
	private lastRow = 0;
	readonly rows = new Map<number, Row>();
	readonly columnDefinitions = new Map<number, Column>();
	readonly merges: string[] = [];
	views?: WorksheetOptions['views'];
	pageSetup: PageSetup;
	headerFooter: { oddFooter?: string } = {};
	autoFilter?: { from: string | Address; to: string | Address };
	constructor(
		readonly name: string,
		options: WorksheetOptions = {}
	) {
		this.views = options.views;
		this.pageSetup = options.pageSetup ?? {};
	}
	get rowCount(): number {
		return this.lastRow;
	}
	get columnCount(): number {
		let count = 0;
		for (const row of this.rows.values())
			count = Math.max(count, row.cellCount);
		return count;
	}
	getRow(number: number): Row {
		validateRow(number);
		let row = this.rows.get(number);
		if (!row) {
			row = new Row(this, number);
			this.rows.set(number, row);
			this.lastRow = Math.max(this.lastRow, number);
		}
		return row;
	}
	addRow(values: unknown[]): Row {
		const row = this.getRow(this.rowCount + 1);
		row.values = values;
		return row;
	}
	getCell(address: string): Cell;
	getCell(row: number, column: number): Cell;
	getCell(address: string | number, column?: number): Cell {
		const position =
			typeof address === 'string'
				? parseAddress(address)
				: { row: address, column: column! };
		return this.getRow(position.row).getCell(position.column);
	}
	getColumn(number: number): Column {
		let column = this.columnDefinitions.get(number);
		if (!column) {
			column = new Column(number);
			this.columnDefinitions.set(number, column);
		}
		return column;
	}
	set columns(columns: Array<{ width: number }>) {
		columns.forEach(
			(column, index) => (this.getColumn(index + 1).width = column.width)
		);
	}
	eachRow(callback: (row: Row, number: number) => void) {
		for (const [number, row] of [...this.rows].sort(([a], [b]) => a - b))
			if (row.values.length > 0) callback(row, number);
	}
	mergeCells(range: string): void;
	mergeCells(top: number, left: number, bottom: number, right: number): void;
	mergeCells(
		range: string | number,
		left?: number,
		bottom?: number,
		right?: number
	): void {
		const [start, end] =
			typeof range === 'string'
				? range.split(':').map(parseAddress)
				: [
						{ row: range, column: left! },
						{ row: bottom!, column: right! }
					];
		if (!end || start.row > end.row || start.column > end.column)
			throw new XlsxError('Invalid merge range.');
		const master = this.getCell(start.row, start.column);
		for (let row = start.row; row <= end.row; row++)
			for (let column = start.column; column <= end.column; column++) {
				const cell = this.getCell(row, column);
				if (
					cell.master ||
					(cell === master &&
						this.merges.some((merge) => merge.startsWith(master.address + ':')))
				)
					throw new XlsxError('Overlapping merge ranges.');
				if (cell !== master) cell.master = master;
			}
		this.merges.push(`${master.address}:${cellAddress(end.row, end.column)}`);
	}
}

export class Workbook {
	readonly worksheets: Worksheet[] = [];
	creator = 'MinkExcel';
	created = new Date();
	modified = this.created;
	calcProperties = { fullCalcOnLoad: true };
	addWorksheet(name: string, options?: WorksheetOptions): Worksheet {
		if (
			!name ||
			name.length > 31 ||
			/[\\/*?:\[\]]/.test(name) ||
			name.startsWith("'") ||
			name.endsWith("'") ||
			this.worksheets.some(
				(sheet) => sheet.name.toLowerCase() === name.toLowerCase()
			)
		)
			throw new XlsxError('Invalid or duplicate worksheet name.');
		const sheet = new Worksheet(name, options);
		this.worksheets.push(sheet);
		return sheet;
	}
	getWorksheet(name: string): Worksheet | undefined {
		return this.worksheets.find((sheet) => sheet.name === name);
	}
}
