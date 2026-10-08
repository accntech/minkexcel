import { expect, test } from 'bun:test';
import { parseXml, stringText, excelText, attributes } from '../src/xml.js';

test('writes uppercase Excel control escapes for readers that require canonical hex digits', () => {
	expect(excelText('\r\u001F')).toBe('_x000D__x001F_');
});

test('serializes only own XML attributes, excluding inherited properties', () => {
	const values = Object.assign(Object.create({ inherited: 'unexpected' }), { id: 1 });
	expect(attributes(values)).toBe(' id="1"');
});

test('visits completed elements and releases consumed rows without retaining the worksheet tree', () => {
	const rows: string[] = [];
	const root = parseXml('<worksheet><sheetData><row><c><v>123</v></c></row><row/></sheetData></worksheet>', (node, parent, depth) => {
		if (node.name === 'row' && parent?.name === 'sheetData' && depth === 2) {
			rows.push(node.children[0]?.children[0]?.text ?? '');
			return true;
		}
		return false;
	});
	expect(rows).toEqual(['123', '']);
	expect(root.children[0].children).toEqual([]);
});

test('reads prefixed elements, quoted attributes, entities, comments and CDATA', () => {
	const node = parseXml(
		'<?xml version="1.0"?><s:si xmlns:s="urn:sheet" label="Ana &amp; &#x1F9FE;"><!-- comment --><s:r><s:t><![CDATA[<Co>]]></s:t></s:r><s:t> &#65;&#10;</s:t></s:si>'
	);
	expect(node.name).toBe('si');
	expect(node.attributes.label).toBe('Ana & 🧾');
	expect(stringText(node)).toBe('<Co> A\n');
});
for (const source of [
	'<a><b></a>',
	'<a/>trailing',
	'<a/><b/>',
	'<a x="1" x="2"/>',
	'<a>&unknown;</a>',
	'<a>&#0;</a>',
	'<a x="<"/>',
	'<a><!--bad--comment--></a>',
	'<!DOCTYPE a><a/>',
	'<a>\u0001</a>',
	'<p:a/>',
	'<a xmlns:p="urn:same" xmlns:q="urn:same" p:id="1" q:id="2"/>',
	'<a xmlns:xml="urn:wrong" xml:space="preserve"/>',
	'<a xmlns:p="http://www.w3.org/XML/1998/namespace"/>',
	'<a xmlns:p="http://www.w3.org/2000/xmlns/"/>',
	'<a xmlns:xmlns="urn:wrong"/>',
	'<xmlns:a/>',
	'<a xmlns="http://www.w3.org/XML/1998/namespace"/>',
	'<a xmlns:p=""/>',
	'<constructor:a/>',
	'<a>\uD800</a>',
	'<a>\uDC00</a>',
	'<a><![CDATA[unfinished</a>'
]) {
	test(`rejects malformed XML ${JSON.stringify(source)}`, () =>
		expect(() => parseXml(source)).toThrow());
}
test('rejects excessive XML nesting', () =>
	expect(() => parseXml('<a>'.repeat(66) + '</a>'.repeat(66))).toThrow('deep'));
