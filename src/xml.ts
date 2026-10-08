import { XlsxError } from './model.js';

export const spreadsheetNamespace = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
export const relationshipNamespace =
	'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
export const declaration = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';

export function xml(value: string | number | boolean): string {
	return String(value)
		.replace(/&/g, '&amp;')
		.replace(/</g, '&lt;')
		.replace(/>/g, '&gt;')
		.replace(/"/g, '&quot;')
		.replace(/'/g, '&apos;');
}
export function excelText(value: string): string {
	return xml(
		value
			.replace(/_x[\da-f]{4}_/gi, (match) => '_x005F_' + match.slice(1))
			.replace(
				/[\u0000-\u0008\u000B\u000C\u000D\u000E-\u001F\uFFFE\uFFFF]/g,
				(character) => `_x${character.charCodeAt(0).toString(16).padStart(4, '0')}_`
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
	return Object.entries(values)
		.filter(([, value]) => value !== undefined)
		.map(
			([name, value]) => ` ${name}="${xml(typeof value === 'boolean' ? Number(value) : value!)}"`
		)
		.join('');
}

export type XmlNode = {
	name: string;
	attributes: Record<string, string>;
	text: string;
	children: XmlNode[];
};
function invalidXml(): never {
	throw new XlsxError('Invalid workbook XML.');
}
function validCharacters(value: string): boolean {
	for (const character of value) {
		const code = character.codePointAt(0)!;
		if (
			(code < 32 && code !== 9 && code !== 10 && code !== 13) ||
			(code >= 0xd800 && code <= 0xdfff) ||
			code === 0xfffe ||
			code === 0xffff
		)
			return false;
	}
	return true;
}
function entities(value: string): string {
	return value.replace(/&([^&;<]*);|&/g, (match, entity: string | undefined) => {
		const named: Record<string, string> = { amp: '&', lt: '<', gt: '>', quot: '"', apos: "'" };
		if (entity && Object.hasOwn(named, entity)) return named[entity];
		if (!entity || !/^#(?:[0-9]+|x[0-9a-f]+)$/i.test(entity)) return invalidXml();
		const code =
			entity[1].toLowerCase() === 'x' ? parseInt(entity.slice(2), 16) : Number(entity.slice(1));
		if (code > 0x10ffff || !validCharacters(String.fromCodePoint(code))) return invalidXml();
		return String.fromCodePoint(code);
	});
}
/** Focused XML 1.0 tokenizer. Never resolves DTDs or external entities. */
export function parseXml(input: string): XmlNode {
	if (!validCharacters(input)) invalidXml();
	const source = input.replace(/^\uFEFF/, '').replace(/\r\n?/g, '\n');
	const stack: Array<{ node: XmlNode; qualified: string; namespaces: Record<string, string> }> = [];
	let root: XmlNode | undefined,
		offset = 0;
	const namePattern = /[A-Za-z_][A-Za-z0-9_.-]*(?::[A-Za-z_][A-Za-z0-9_.-]*)?/y;
	const name = () => {
		namePattern.lastIndex = offset;
		const match = namePattern.exec(source);
		if (!match) return invalidXml();
		offset = namePattern.lastIndex;
		return match[0];
	};
	const whitespace = () => {
		const start = offset;
		while (/[\t\n\r ]/.test(source[offset] ?? '') && offset < source.length) offset++;
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
			stack.pop();
			continue;
		}
		const qualified = name(),
			attrs: Record<string, string> = Object.create(null);
		for (;;) {
			const separated = whitespace();
			if (source[offset] === '>' || source.startsWith('/>', offset)) break;
			if (!separated) invalidXml();
			const attribute = name();
			if (Object.hasOwn(attrs, attribute)) invalidXml();
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
		const namespaces = {
			...(stack.at(-1)?.namespaces ?? { xml: 'http://www.w3.org/XML/1998/namespace' })
		};
		for (const [key, value] of Object.entries(attrs))
			if (key.startsWith('xmlns:')) namespaces[key.slice(6)] = value;
		const checkPrefix = (key: string) => {
			if (key.includes(':') && !key.startsWith('xmlns:') && !namespaces[key.split(':')[0]])
				invalidXml();
		};
		checkPrefix(qualified);
		for (const key of Object.keys(attrs)) checkPrefix(key);
		const node: XmlNode = {
			name: qualified.split(':').at(-1)!,
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
		if (source.startsWith('/>', offset)) offset += 2;
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
	return decodeExcelText(
		node.children
			.filter((entry) => entry.name === 't' || entry.name === 'r')
			.map((entry) =>
				entry.name === 't'
					? entry.text
					: entry.children
							.filter((part) => part.name === 't')
							.map((part) => part.text)
							.join('')
			)
			.join('')
	);
}
