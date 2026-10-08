import { expect, test } from 'bun:test';
import { parseXml, stringText } from '../src/xml.js';

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
	'<a><![CDATA[unfinished</a>'
]) {
	test(`rejects malformed XML ${JSON.stringify(source)}`, () =>
		expect(() => parseXml(source)).toThrow());
}
test('rejects excessive XML nesting', () =>
	expect(() => parseXml('<a>'.repeat(66) + '</a>'.repeat(66))).toThrow('deep'));
