import { XlsxError } from './model.js';

export const spreadsheetNamespace = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
export const relationshipNamespace =
	'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
export const declaration = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
const xmlEscapes: Record<string, string> = { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&apos;' };

export function xml(value: string | number | boolean): string {
	return String(value).replace(/[&<>"']/g, (character) => xmlEscapes[character]);
}
export function excelText(value: string): string {
	return xml(
		value
			.replace(/_x[\da-f]{4}_/gi, (match) => '_x005F_' + match.slice(1))
			.replace(
				/[\u0000-\u0008\u000B\u000C\u000D\u000E-\u001F\uFFFE\uFFFF]/g,
				(character) => `_x${character.charCodeAt(0).toString(16).toUpperCase().padStart(4, '0')}_`
			)
	);
}
export function decodeExcelText(value: string): string {
	return value.replace(
		/_x005f_(_?x[\da-f]{4}_)|_x([\da-f]{4})_/gi,
		(_match, escaped: string | undefined, code: string | undefined) =>
			escaped ? '_' + escaped.replace(/^_/, '') : String.fromCharCode(parseInt(code!, 16))
	);
}
export function attributes(values: Record<string, string | number | boolean | undefined>): string {
	let result = '';
	for (const name in values) {
		if (!Object.hasOwn(values, name)) continue;
		const value = values[name];
		if (value !== undefined)
			result += ` ${name}="${xml(typeof value === 'boolean' ? Number(value) : value)}"`;
	}
	return result;
}

export type XmlNode = {
	name: string;
	attributes: Record<string, string>;
	text: string;
	children: XmlNode[];
	source?: { start: number; end: number; openEnd: number; closeStart: number; qualified: string };
};
export type XmlVisitor = (node: XmlNode, parent: XmlNode | undefined, depth: number) => boolean;
function invalidXml(): never {
	throw new XlsxError('Invalid workbook XML.');
}
function validCharacters(value: string): boolean {
	// Unicode mode rejects lone surrogates while permitting valid astral pairs.
	return !/[\u0000-\u0008\u000B\u000C\u000E-\u001F\uD800-\uDFFF\uFFFE\uFFFF]/u.test(value);
}
const namedEntities: Record<string, string> = { amp: '&', lt: '<', gt: '>', quot: '"', apos: "'" };
function checkPrefix(key: string, namespaces: Record<string, string>): void {
	const colon = key.indexOf(':');
	if (colon >= 0 && !namespaces[key.slice(0, colon)]) invalidXml();
}
function entities(value: string): string {
	if (!value.includes('&')) return value;
	return value.replace(/&([^&;<]*);|&/g, (match, entity: string | undefined) => {
		if (entity && Object.hasOwn(namedEntities, entity)) return namedEntities[entity];
		if (!entity || !/^#(?:[0-9]+|x[0-9a-f]+)$/i.test(entity)) return invalidXml();
		const code =
			entity[1].toLowerCase() === 'x' ? parseInt(entity.slice(2), 16) : Number(entity.slice(1));
		if (code > 0x10ffff || !validCharacters(String.fromCodePoint(code))) return invalidXml();
		return String.fromCodePoint(code);
	});
}
const xmlNamespace = 'http://www.w3.org/XML/1998/namespace';
const xmlnsNamespace = 'http://www.w3.org/2000/xmlns/';
type OpenElement = { node: XmlNode; qualified: string; namespaces: Record<string, string> };

function letter(code: number): boolean {
	return (code >= 65 && code <= 90) || (code >= 97 && code <= 122);
}
function nameStart(code: number): boolean {
	return code === 95 || letter(code);
}
function nameCharacter(code: number): boolean {
	return nameStart(code) || code === 45 || code === 46 || (code >= 48 && code <= 57);
}
function xmlWhitespace(code: number): boolean {
	return code === 32 || code === 9 || code === 10 || code === 13;
}
function validateNamespace(key: string, value: string): void {
	if (key === 'xmlns') {
		if (value === xmlNamespace || value === xmlnsNamespace) invalidXml();
		return;
	}
	if (key === 'xmlns:xml') {
		if (value !== xmlNamespace) invalidXml();
		return;
	}
	if (!value || key === 'xmlns:xmlns' || value === xmlnsNamespace || value === xmlNamespace)
		invalidXml();
}
function elementNamespaces(attrs: Record<string, string>, keys: string[] | undefined, inherited: Record<string, string>): Record<string, string> {
	if (!keys) return inherited;
	let namespaces = inherited;
	let copied = false;
	for (const key of keys) {
		validateNamespace(key, attrs[key]);
		if (key === 'xmlns') continue;
		if (!copied) {
			namespaces = Object.assign(Object.create(null), inherited);
			copied = true;
		}
		namespaces[key.slice(6)] = attrs[key];
	}
	return namespaces;
}
function validateAttributes(keys: string[] | undefined, namespaces: Record<string, string>): void {
	if (!keys) return;
	const expandedAttributes = new Set<string>();
	for (const key of keys) {
		checkPrefix(key, namespaces);
		const colon = key.indexOf(':');
		const expanded = namespaces[key.slice(0, colon)] + '\u0000' + key.slice(colon + 1);
		if (expandedAttributes.has(expanded)) invalidXml();
		expandedAttributes.add(expanded);
	}
}

class XmlParser {
	private offset = 0;
	private stack: OpenElement[] = [];
	private root?: XmlNode;
	private rootNamespaces: Record<string, string> = Object.assign(Object.create(null), { xml: xmlNamespace });
	constructor(private source: string, private onClose?: XmlVisitor, private captureSource = false) {}

	parse(): XmlNode {
		while (this.offset < this.source.length) this.token();
		if (this.stack.length) invalidXml();
		if (!this.root) throw new XlsxError('Empty workbook XML.');
		return this.root;
	}
	private token(): void {
		if (this.source[this.offset] !== '<') return this.text();
		if (this.source.startsWith('<!--', this.offset)) return this.comment();
		if (this.source.startsWith('<![CDATA[', this.offset)) return this.cdata();
		if (this.source.startsWith('<?', this.offset)) return this.instruction();
		if (this.source.startsWith('<!', this.offset))
			throw new XlsxError('Workbook XML must not contain a DTD.');
		const start = this.offset++;
		if (this.source[this.offset] === '/') return this.closeElement(start);
		this.openElement(start);
	}
	private complete(node: XmlNode): void {
		if (!this.onClose) return;
		const parent = this.stack.at(-1)?.node;
		if (this.onClose(node, parent, this.stack.length) && parent) parent.children.pop();
	}
	private name(): string {
		const start = this.offset;
		for (let part = 0; part < 2; part++) {
			if (!nameStart(this.source.charCodeAt(this.offset))) invalidXml();
			this.offset++;
			while (this.offset < this.source.length && nameCharacter(this.source.charCodeAt(this.offset)))
				this.offset++;
			if (this.source.charCodeAt(this.offset) !== 58) break;
			if (part === 1) invalidXml();
			this.offset++;
		}
		return this.source.slice(start, this.offset);
	}
	private whitespace(): boolean {
		const start = this.offset;
		while (this.offset < this.source.length && xmlWhitespace(this.source.charCodeAt(this.offset)))
			this.offset++;
		return this.offset > start;
	}
	private text(): void {
		const next = this.source.indexOf('<', this.offset);
		const end = next < 0 ? this.source.length : next;
		const raw = this.source.slice(this.offset, end);
		if (raw.includes(']]>')) invalidXml();
		const text = entities(raw);
		const parent = this.stack.at(-1);
		if (parent) parent.node.text += text;
		else if (text.trim()) invalidXml();
		this.offset = end;
	}
	private comment(): void {
		const end = this.source.indexOf('-->', this.offset + 4);
		if (end < 0 || this.source.slice(this.offset + 4, end).includes('--') || this.source[end - 1] === '-')
			invalidXml();
		this.offset = end + 3;
	}
	private cdata(): void {
		const end = this.source.indexOf(']]>', this.offset + 9);
		const parent = this.stack.at(-1);
		if (end < 0 || !parent) invalidXml();
		parent.node.text += this.source.slice(this.offset + 9, end);
		this.offset = end + 3;
	}
	private instruction(): void {
		const start = this.offset;
		const end = this.source.indexOf('?>', this.offset + 2);
		if (end < 0) invalidXml();
		this.offset += 2;
		const target = this.name();
		if (!this.whitespace() && this.offset !== end) invalidXml();
		if (target.toLowerCase() === 'xml') this.declaration(target, start, end);
		this.offset = end + 2;
	}
	private declaration(target: string, start: number, end: number): void {
		if (target !== 'xml' || start !== 0) invalidXml();
		const content = this.source.slice(this.offset, end);
		if (!/^version\s*=\s*(["'])1\.0\1(?:\s+encoding\s*=\s*(["'])UTF-8\2)?(?:\s+standalone\s*=\s*(["'])(?:yes|no)\3)?\s*$/i.test(content))
			invalidXml();
	}
	private closeElement(start: number): void {
		this.offset++;
		const qualified = this.name();
		this.whitespace();
		if (this.source[this.offset++] !== '>' || this.stack.at(-1)?.qualified !== qualified)
			invalidXml();
		const node = this.stack.pop()!.node;
		if (node.source) {
			node.source.closeStart = start;
			node.source.end = this.offset;
		}
		this.complete(node);
	}
	private attributeValue(): string {
		this.whitespace();
		if (this.source[this.offset++] !== '=') invalidXml();
		this.whitespace();
		const quote = this.source[this.offset++];
		if (quote !== '"' && quote !== "'") invalidXml();
		const end = this.source.indexOf(quote, this.offset);
		if (end < 0) invalidXml();
		const value = this.source.slice(this.offset, end);
		if (value.includes('<')) invalidXml();
		this.offset = end + 1;
		return entities(value.replace(/[\t\n]/g, ' '));
	}
	private elementAttributes() {
		const attrs: Record<string, string> = Object.create(null);
		let namespaces: string[] | undefined;
		let prefixed: string[] | undefined;
		for (;;) {
			const separated = this.whitespace();
			if (this.source[this.offset] === '>' || this.source.startsWith('/>', this.offset)) break;
			if (!separated) invalidXml();
			const name = this.name();
			if (Object.hasOwn(attrs, name)) invalidXml();
			if (name === 'xmlns' || name.startsWith('xmlns:')) (namespaces ??= []).push(name);
			else if (name.includes(':')) (prefixed ??= []).push(name);
			attrs[name] = this.attributeValue();
		}
		return { attrs, namespaces, prefixed };
	}
	private openElement(start: number): void {
		const qualified = this.name();
		const { attrs, namespaces: keys, prefixed } = this.elementAttributes();
		const inherited = this.stack.at(-1)?.namespaces ?? this.rootNamespaces;
		const namespaces = elementNamespaces(attrs, keys, inherited);
		if (qualified.startsWith('xmlns:')) invalidXml();
		checkPrefix(qualified, namespaces);
		validateAttributes(prefixed, namespaces);
		const node: XmlNode = {
			name: qualified.slice(qualified.indexOf(':') + 1), attributes: attrs, text: '', children: []
		};
		if (this.stack.length >= 64) throw new XlsxError('Workbook XML is nested too deeply.');
		if (this.captureSource)
			node.source = { start, end: 0, openEnd: this.offset + 1, closeStart: 0, qualified };
		this.attach(node);
		if (this.source.startsWith('/>', this.offset)) {
			this.offset += 2;
			if (node.source) {
				node.source.openEnd = node.source.end = this.offset;
				node.source.closeStart = this.offset;
			}
			this.complete(node);
			return;
		}
		this.offset++;
		this.stack.push({ node, qualified, namespaces });
	}
	private attach(node: XmlNode): void {
		const parent = this.stack.at(-1);
		if (parent) parent.node.children.push(node);
		else {
			if (this.root) invalidXml();
			this.root = node;
		}
	}
}
/** Focused XML 1.0 tokenizer. Never resolves DTDs or external entities. */
export function parseXml(input: string, onClose?: XmlVisitor, captureSource = false): XmlNode {
	if (!validCharacters(input)) invalidXml();
	const source = input.replace(/^\uFEFF/, '').replace(/\r\n?/g, '\n');
	return new XmlParser(source, onClose, captureSource).parse();
}
export function children(node: XmlNode, name: string): XmlNode[] {
	return node.children.filter((child) => child.name === name);
}
export function child(node: XmlNode, name: string): XmlNode | undefined {
	return node.children.find((child) => child.name === name);
}
export function stringText(node: XmlNode): string {
	let text = '';
	for (const entry of node.children) {
		if (entry.name === 't') text += entry.text;
		else if (entry.name === 'r') {
			for (const part of entry.children) {
				if (part.name === 't') text += part.text;
			}
		}
	}
	return decodeExcelText(text);
}
