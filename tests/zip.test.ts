import { expect, test } from 'bun:test';
import { ZipArchive, writeZip } from '../src/zip.js';
import { unzip } from './helpers.js';
const limits = {
	fileBytes: 5 * 1024 * 1024,
	entries: 1000,
	entryBytes: 20 * 1024 * 1024,
	totalBytes: 50 * 1024 * 1024,
	compressionRatio: 200
};

test('reads compressed independently generated ZIP entries', async () => {
	const archive = new ZipArchive(
		new Uint8Array(
			await Bun.file(new URL('fixtures/exceljs-1900.xlsx', import.meta.url)).arrayBuffer()
		),
		limits
	);
	expect(new TextDecoder().decode(await archive.read('xl/workbook.xml'))).toContain('<workbook');
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
test('rejects tampered checksums, local headers, duplicate entries and output bounds', async () => {
	const valid = await writeZip(
		new Map([
			['a.xml', 'abc'],
			['b.xml', 'def']
		])
	);
	const crc = valid.slice();
	crc[35] ^= 1;
	await expect(new ZipArchive(crc, limits).read('a.xml')).rejects.toThrow();
	const header = valid.slice();
	header[0] = 0;
	expect(() => new ZipArchive(header, limits)).toThrow();
	const duplicate = valid.slice();
	for (let i = 0; i < duplicate.length - 5; i++)
		if (new TextDecoder().decode(duplicate.subarray(i, i + 5)) === 'b.xml') duplicate[i] = 97;
	expect(() => new ZipArchive(duplicate, limits)).toThrow();
	expect(() => new ZipArchive(valid, { ...limits, entryBytes: 2 })).toThrow('large');
	const central = valid.slice();
	const centralStart = new DataView(central.buffer).getUint32(central.length - 6, true);
	new DataView(central.buffer).setUint32(centralStart + 24, 1, true);
	new DataView(central.buffer).setUint32(22, 1, true);
	await expect(new ZipArchive(central, limits).read('a.xml')).rejects.toThrow('large');
});
test('cancels archive generation before completion', async () => {
	const controller = new AbortController();
	const pending = writeZip(new Map([['large.xml', 'x'.repeat(1000000)]]), controller.signal);
	setTimeout(() => controller.abort(), 0);
	await expect(pending).rejects.toThrow();
});

test('bounds actual DEFLATE output even when the archive understates its size', async () => {
	const bytes = new Uint8Array(
		await Bun.file(new URL('fixtures/exceljs-1900.xlsx', import.meta.url)).arrayBuffer()
	);
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
