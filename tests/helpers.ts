import { inflateRawSync } from 'node:zlib';
import { expect } from 'bun:test';
import { writeZip } from '../src/zip.js';

/** Independent test decoder, using the platform inflater rather than the package reader. */
export function unzip(bytes: Uint8Array): Map<string, string> {
	const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength),
		decoder = new TextDecoder();
	expect(view.getUint32(bytes.length - 22, true)).toBe(0x06054b50);
	const count = view.getUint16(bytes.length - 12, true),
		parts = new Map<string, string>();
	let cursor = view.getUint32(bytes.length - 6, true);
	for (let index = 0; index < count; index++) {
		expect(view.getUint32(cursor, true)).toBe(0x02014b50);
		const method = view.getUint16(cursor + 10, true),
			length = view.getUint32(cursor + 20, true),
			size = view.getUint32(cursor + 24, true),
			nameLength = view.getUint16(cursor + 28, true),
			local = view.getUint32(cursor + 42, true);
		expect(view.getUint32(local, true)).toBe(0x04034b50);
		const start = local + 30 + view.getUint16(local + 26, true) + view.getUint16(local + 28, true),
			data = bytes.subarray(start, start + length);
		const decoded = method === 8 ? inflateRawSync(data) : data;
		expect(decoded.length).toBe(size);
		parts.set(
			decoder.decode(bytes.subarray(cursor + 46, cursor + 46 + nameLength)),
			decoder.decode(decoded)
		);
		cursor +=
			46 + nameLength + view.getUint16(cursor + 30, true) + view.getUint16(cursor + 32, true);
	}
	return parts;
}
export async function fixture(date1904 = false) {
	return new Uint8Array(
		await Bun.file(
			new URL(`fixtures/exceljs-${date1904 ? '1904' : '1900'}.xlsx`, import.meta.url)
		).arrayBuffer()
	);
}

/** Keep the independent producer's package while changing only the parts under test. */
export async function modifiedFixture(parts: Record<string, string | undefined>) {
	const contents = unzip(await fixture());
	for (const [path, xml] of Object.entries(parts)) {
		if (xml === undefined) contents.delete(path);
		else contents.set(path, xml);
	}
	return writeZip(contents);
}
