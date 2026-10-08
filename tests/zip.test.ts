import { expect, test } from 'bun:test';
import { ZipArchive, writeZip, crc32 } from '../src/zip.js';
import { XlsxError, XlsxLimitError } from '../src/index.js';
import { fixture, unzip } from './helpers.js';
const limits = {
	fileBytes: 5 * 1024 * 1024,
	entries: 1000,
	entryBytes: 20 * 1024 * 1024,
	totalBytes: 50 * 1024 * 1024,
	compressionRatio: 200
};

// Python stdlib zipfile, ZIP_STORED, two entries and a ZIP comment.
// An independent archive exercises STORE without reusing our writer.
const stored = Uint8Array.from(Buffer.from(
	'UEsDBBQAAAAAAAAASF3CQSQ1AwAAAAMAAAAFAAAAYS54bWxhYmNQSwMEFAAAAAAAAABIXWHhxAwDAAAAAwAAAAUAAABiLnhtbGRlZlBLAQIUAxQAAAAAAAAASF3CQSQ1AwAAAAMAAAAFAAAAAAAAAAAAAACAAQAAAABhLnhtbFBLAQIUAxQAAAAAAAAASF1h4cQMAwAAAAMAAAAFAAAAAAAAAAAAAACAASYAAABiLnhtbFBLBQYAAAAAAgACAGYAAABMAAAAFwBpbmRlcGVuZGVudCBaSVAgZml4dHVyZQ==',
	'base64'
));

test('reads independent STORE entries with a ZIP comment from a nonzero byte offset', async () => {
	const padded = new Uint8Array(stored.length + 20);
	padded.set(stored, 7);
	const archive = new ZipArchive(padded.subarray(7, 7 + stored.length), limits);
	expect(new TextDecoder().decode(await archive.read('a.xml'))).toBe('abc');
	expect(new TextDecoder().decode(await archive.read('b.xml'))).toBe('def');
});

test('uses the standard CRC32 checksum, including the empty input', () => {
	expect(crc32(new TextEncoder().encode('123456789'))).toBe(0xcbf43926);
	expect(crc32(new Uint8Array())).toBe(0);
});

test('writes UTF-8 filenames and compressed DEFLATE entries', async () => {
	const bytes = await writeZip(new Map([['α.xml', 'Hello & 🧾']]));
	const view = new DataView(bytes.buffer);
	expect(view.getUint32(0, true)).toBe(0x04034b50);
	expect(view.getUint16(8, true)).toBe(8);
	expect(new TextDecoder().decode(await new ZipArchive(bytes, limits).read('α.xml'))).toBe(
		'Hello & 🧾'
	);
});
test('rejects a checksum mismatch after successfully decompressing intact data', async () => {
	const bytes = await writeZip(new Map([['a.xml', 'abc']]));
	const view = new DataView(bytes.buffer);
	const central = view.getUint32(bytes.length - 6, true);
	const wrong = view.getUint32(14, true) ^ 1;
	view.setUint32(14, wrong, true);
	view.setUint32(central + 16, wrong, true);
	await expect(new ZipArchive(bytes, limits).read('a.xml')).rejects.toThrow('checksum');
});

for (const [label, mutate] of [
	['local header signature', (view: DataView, central: number) => view.setUint32(0, 0, true)],
	['central header signature', (view: DataView, central: number) => view.setUint32(central, 0, true)],
	['mismatched compression method', (view: DataView, central: number) => view.setUint16(8, 0, true)],
	['mismatched flags', (view: DataView, central: number) => view.setUint16(6, 0, true)],
	['mismatched local filename', (view: DataView, central: number) => view.setUint8(30, 98)],
	['mismatched local size', (view: DataView, central: number) => view.setUint32(22, 99, true)],
	['encrypted entry', (view: DataView, central: number) => {
		view.setUint16(6, 0x801, true);
		view.setUint16(central + 8, 0x801, true);
	}],
	['unsupported compression method', (view: DataView, central: number) => {
		view.setUint16(8, 99, true);
		view.setUint16(central + 10, 99, true);
	}],
	['entry extending into the central directory', (view: DataView, central: number) => {
		view.setUint32(18, central, true);
		view.setUint32(central + 20, central, true);
	}],
	['multi-disk archive', (view: DataView, central: number) => view.setUint16(view.byteLength - 18, 1, true)],
	['ZIP64 entry count', (view: DataView, central: number) => {
		view.setUint16(view.byteLength - 14, 65535, true);
		view.setUint16(view.byteLength - 12, 65535, true);
	}]
] as const) {
	test(`rejects ${label}`, async () => {
		const bytes = await writeZip(new Map([['a.xml', 'abc']]));
		const view = new DataView(bytes.buffer);
		mutate(view, view.getUint32(bytes.length - 6, true));
		expect(() => new ZipArchive(bytes, limits)).toThrow(XlsxError);
	});
}

test('rejects duplicate entry names in otherwise consistent local and central headers', async () => {
	const bytes = await writeZip(new Map([['a.xml', 'abc'], ['b.xml', 'def']]));
	const view = new DataView(bytes.buffer);
	const first = view.getUint32(bytes.length - 6, true);
	const second = first + 46 + view.getUint16(first + 28, true);
	const local = view.getUint32(second + 42, true);
	bytes[second + 46] = 97;
	bytes[local + 30] = 97;
	expect(() => new ZipArchive(bytes, limits)).toThrow(XlsxError);
});

test('accepts exact archive limits and rejects tighter bounds with typed limit errors', async () => {
	const exact = { ...limits, fileBytes: stored.length, entries: 2, entryBytes: 3, totalBytes: 6, compressionRatio: 1 };
	expect(await new ZipArchive(stored, exact).read('a.xml')).toEqual(new TextEncoder().encode('abc'));
	for (const tighter of [{ fileBytes: stored.length - 1 }, { entries: 1 }, { entryBytes: 2 }, { totalBytes: 5 }, { compressionRatio: 0.5 }])
		expect(() => new ZipArchive(stored, { ...exact, ...tighter })).toThrow(XlsxLimitError);
	const archive = new ZipArchive(stored, exact);
	await expect(archive.read('a.xml', 2)).rejects.toThrow(XlsxLimitError);
	await expect(archive.read('missing.xml')).rejects.toThrow('Missing workbook part: missing.xml');
});

test('rejects truncated archives and unexpected trailing bytes', () => {
	for (const bytes of [new Uint8Array(), stored.subarray(0, 21), stored.subarray(0, stored.length - 1), new Uint8Array([...stored, 0])])
		expect(() => new ZipArchive(bytes, limits)).toThrow(XlsxError);
});

for (const path of ['', '/a.xml', '../a.xml', 'xl/../a.xml', 'xl/./a.xml', 'xl\\a.xml', 'C:a.xml', 'a\u0000.xml']) {
	test(`rejects unsafe entry path ${JSON.stringify(path)}`, async () => {
		await expect(writeZip(new Map([[path, 'abc']]))).rejects.toThrow(XlsxError);
	});
}

test('bounds actual DEFLATE output even when the archive understates its size', async () => {
	const bytes = await fixture();
	const view = new DataView(bytes.buffer);
	let cursor = view.getUint32(bytes.length - 6, true);
	for (let i = 0; i < view.getUint16(bytes.length - 12, true); i++) {
		const length = view.getUint16(cursor + 28, true),
			name = new TextDecoder().decode(bytes.subarray(cursor + 46, cursor + 46 + length));
		if (name === 'xl/workbook.xml') {
			const local = view.getUint32(cursor + 42, true);
			view.setUint32(cursor + 24, 1, true);
			view.setUint32(local + 22, 1, true);
			break;
		}
		cursor += 46 + length + view.getUint16(cursor + 30, true) + view.getUint16(cursor + 32, true);
	}
	await expect(new ZipArchive(bytes, limits).read('xl/workbook.xml')).rejects.toThrow('large');
});

test('compresses repetitive XML and writes correct sizes in both ZIP headers', async () => {
	const xml = '<row><c t="inlineStr"><is><t>Customer</t></is></c></row>'.repeat(1000);
	const bytes = await writeZip(new Map([['sheet.xml', xml]])),
		view = new DataView(bytes.buffer),
		central = view.getUint32(bytes.length - 6, true);
	expect(bytes.length).toBeLessThan(new TextEncoder().encode(xml).length / 10);
	expect(view.getUint16(central + 10, true)).toBe(8);
	expect(view.getUint32(18, true)).toBe(view.getUint32(central + 20, true));
	expect(view.getUint32(22, true)).toBe(new TextEncoder().encode(xml).length);
	expect(view.getUint32(central + 24, true)).toBe(new TextEncoder().encode(xml).length);
	expect(unzip(bytes).get('sheet.xml')).toBe(xml);
});
