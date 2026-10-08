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
/** Focused XML 1.0 tokenizer. Never resolves DTDs or external entities. */
export function parseXml(input: string, onClose?: XmlVisitor): XmlNode {
	if (!validCharacters(input)) invalidXml();
	const source = input.replace(/^\uFEFF/, '').replace(/\r\n?/g, '\n');
	const stack: Array<{ node: XmlNode; qualified: string; namespaces: Record<string, string> }> = [];
	const rootNamespaces: Record<string, string> = Object.assign(Object.create(null), {
		xml: 'http://www.w3.org/XML/1998/namespace'
	});
	let root: XmlNode | undefined,
		offset = 0;
	const complete = (node: XmlNode) => {
		if (!onClose) return;
		const parent = stack.at(-1)?.node;
		if (onClose(node, parent, stack.length) && parent) parent.children.pop();
	};
	const name = () => {
		const start = offset;
		for (let part = 0; part < 2; part++) {
			let code = source.charCodeAt(offset);
			if (!(code === 95 || code >= 65 && code <= 90 || code >= 97 && code <= 122)) invalidXml();
			offset++;
			while (offset < source.length) {
				code = source.charCodeAt(offset);
				if (!(code === 95 || code === 45 || code === 46 || code >= 48 && code <= 57 ||
					code >= 65 && code <= 90 || code >= 97 && code <= 122)) break;
				offset++;
			}
			if (source.charCodeAt(offset) !== 58) break;
			if (part === 1) invalidXml();
			offset++;
		}
		return source.slice(start, offset);
	};
	const whitespace = () => {
		const start = offset;
		while (offset < source.length) {
			const code = source.charCodeAt(offset);
			if (code !== 32 && code !== 9 && code !== 10 && code !== 13) break;
			offset++;
		}
		return offset > start;
	};
	while (offset < source.length) {
		if (source[offset] !== '<') {
			const next = source.indexOf('<', offset),
				end = next < 0 ? source.length : next;
			const raw = source.slice(offset, end);
			if (raw.includes(']]>')) invalidXml();
			const text = entities(raw);
			if (stack.length) stack.at(-1)!.node.text += text;
			else if (text.trim()) invalidXml();
			offset = end;
			continue;
		}
		if (source.startsWith('<!--', offset)) {
			const end = source.indexOf('-->', offset + 4);
			if (end < 0 || source.slice(offset + 4, end).includes('--') || source[end - 1] === '-')
				invalidXml();
			offset = end + 3;
			continue;
		}
		if (source.startsWith('<![CDATA[', offset)) {
			const end = source.indexOf(']]>', offset + 9);
			if (end < 0 || !stack.length) invalidXml();
			stack.at(-1)!.node.text += source.slice(offset + 9, end);
			offset = end + 3;
			continue;
		}
		if (source.startsWith('<?', offset)) {
			const start = offset,
				end = source.indexOf('?>', offset + 2);
			if (end < 0) invalidXml();
			offset += 2;
			const target = name();
			if (!whitespace() && offset !== end) invalidXml();
			if (target.toLowerCase() === 'xml') {
				if (target !== 'xml' || start !== 0) invalidXml();
				const declaration = source.slice(offset, end);
				if (
					!/^version\s*=\s*(["'])1\.0\1(?:\s+encoding\s*=\s*(["'])UTF-8\2)?(?:\s+standalone\s*=\s*(["'])(?:yes|no)\3)?\s*$/i.test(
						declaration
					)
				)
					invalidXml();
			}
			offset = end + 2;
			continue;
		}
		if (source.startsWith('<!', offset))
			throw new XlsxError('Workbook XML must not contain a DTD.');
		offset++;
		if (source[offset] === '/') {
			offset++;
			const qualified = name();
			whitespace();
			if (source[offset++] !== '>' || !stack.length || stack.at(-1)!.qualified !== qualified)
				invalidXml();
			complete(stack.pop()!.node);
			continue;
		}
		const qualified = name(),
			attrs: Record<string, string> = Object.create(null);
		let namespaceAttributes: string[] | undefined;
		let prefixedAttributes: string[] | undefined;
		for (;;) {
			const separated = whitespace();
			if (source[offset] === '>' || source.startsWith('/>', offset)) break;
			if (!separated) invalidXml();
			const attribute = name();
			if (Object.hasOwn(attrs, attribute)) invalidXml();
			if (attribute === 'xmlns' || attribute.startsWith('xmlns:'))
				(namespaceAttributes ??= []).push(attribute);
			else if (attribute.includes(':')) (prefixedAttributes ??= []).push(attribute);
			whitespace();
			if (source[offset++] !== '=') invalidXml();
			whitespace();
			const quote = source[offset++];
			if (quote !== '"' && quote !== "'") invalidXml();
			const end = source.indexOf(quote, offset);
			if (end < 0) invalidXml();
			const value = source.slice(offset, end);
			if (value.includes('<')) invalidXml();
			attrs[attribute] = entities(value.replace(/[\t\n]/g, ' '));
			offset = end + 1;
		}
		let namespaces = stack.at(-1)?.namespaces ?? rootNamespaces;
		let declaredNamespaces = false;
		for (const key of namespaceAttributes ?? []) {
			if (key === 'xmlns' && (attrs[key] === 'http://www.w3.org/XML/1998/namespace' || attrs[key] === 'http://www.w3.org/2000/xmlns/')) invalidXml();
			if (key.startsWith('xmlns:')) {
				const value = attrs[key];
				if (key === 'xmlns:xml' && value !== 'http://www.w3.org/XML/1998/namespace') invalidXml();
				if (!value || key === 'xmlns:xmlns' || value === 'http://www.w3.org/2000/xmlns/' ||
					(key !== 'xmlns:xml' && value === 'http://www.w3.org/XML/1998/namespace')) invalidXml();
				if (!declaredNamespaces) {
					namespaces = Object.assign(Object.create(null), namespaces);
					declaredNamespaces = true;
				}
				namespaces[key.slice(6)] = value;
			}
		}
		if (qualified.startsWith('xmlns:')) invalidXml();
		checkPrefix(qualified, namespaces);
		let expandedAttributes: Set<string> | undefined;
		for (const key of prefixedAttributes ?? []) {
			const colon = key.indexOf(':');
			checkPrefix(key, namespaces);
			expandedAttributes ??= new Set<string>();
			const expanded = namespaces[key.slice(0, colon)] + '\u0000' + key.slice(colon + 1);
			if (expandedAttributes.has(expanded)) invalidXml();
			expandedAttributes.add(expanded);
		}
		const node: XmlNode = {
			name: qualified.slice(qualified.indexOf(':') + 1),
			attributes: attrs,
			text: '',
			children: []
		};
		if (stack.length >= 64) throw new XlsxError('Workbook XML is nested too deeply.');
		if (stack.length) stack.at(-1)!.node.children.push(node);
		else {
			if (root) invalidXml();
			root = node;
		}
		if (source.startsWith('/>', offset)) {
			offset += 2;
			complete(node);
		}
		else {
			offset++;
			stack.push({ node, qualified, namespaces });
		}
	}
	if (stack.length) invalidXml();
	if (!root) throw new XlsxError('Empty workbook XML.');
	return root;
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
