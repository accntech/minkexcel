import { XlsxError, XlsxLimitError } from './model.js';
import { inflate } from './inflate.js';
import { deflate } from './deflate.js';
import { checkpoint } from './async.js';
export type ArchiveLimits = {
	fileBytes: number;
	entries: number;
	entryBytes: number;
	totalBytes: number;
	compressionRatio: number;
};
type Entry = {
	name: string;
	method: number;
	crc: number;
	size: number;
	start: number;
	end: number;
	local: number;
};
// These local initializations have no externally visible effects when unused.
const encoder = /* @__PURE__ */ new TextEncoder();
const decoder = /* @__PURE__ */ new TextDecoder('utf-8', { fatal: true });
const crcTable = /* @__PURE__ */ Uint32Array.from({ length: 256 }, (_, i) => {
	let value = i;
	for (let bit = 0; bit < 8; bit++) value = value & 1 ? 0xedb88320 ^ (value >>> 1) : value >>> 1;
	return value >>> 0;
});
export function crc32(bytes: Uint8Array): number {
	let crc = 0xffffffff;
	for (const byte of bytes) crc = crcTable[(crc ^ byte) & 255] ^ (crc >>> 8);
	return (crc ^ 0xffffffff) >>> 0;
}
function invalid(): never {
	throw new XlsxError('The file is not a valid .xlsx workbook archive.');
}
function entryName(bytes: Uint8Array, flags: number): string {
	if (!(flags & 0x800) && bytes.some((byte) => byte >= 128)) invalid();
	const name = decoder.decode(bytes);
	if (
		!name ||
		name.startsWith('/') ||
		/[\\:\u0000]/.test(name) ||
		name.split('/').some((part) => part === '..' || part === '.')
	)
		invalid();
	return name;
}
type Directory = { end: number; central: number; count: number };
type EntryHeader = {
	flags: number;
	method: number;
	crc: number;
	compressed: number;
	size: number;
	local: number;
	nameLength: number;
	next: number;
};
function endRecord(view: DataView): number {
	for (let offset = view.byteLength - 22; offset >= Math.max(0, view.byteLength - 22 - 65535); offset--) {
		if (view.getUint32(offset, true) === 0x06054b50 &&
			offset + 22 + view.getUint16(offset + 20, true) === view.byteLength)
			return offset;
	}
	return invalid();
}
function directory(view: DataView, limits: ArchiveLimits): Directory {
	const end = endRecord(view);
	const count = view.getUint16(end + 10, true);
	const size = view.getUint32(end + 12, true);
	const central = view.getUint32(end + 16, true);
	if (view.getUint16(end + 4, true) || view.getUint16(end + 6, true) ||
		view.getUint16(end + 8, true) !== count || count === 65535 ||
		size === 0xffffffff || central === 0xffffffff || central + size !== end)
		invalid();
	if (count > limits.entries) throw new XlsxLimitError('The workbook contains too many archive entries.');
	return { end, central, count };
}
function centralHeader(view: DataView, cursor: number, end: number): EntryHeader {
	if (cursor + 46 > end || view.getUint32(cursor, true) !== 0x02014b50) invalid();
	const nameLength = view.getUint16(cursor + 28, true);
	const header: EntryHeader = {
		flags: view.getUint16(cursor + 8, true),
		method: view.getUint16(cursor + 10, true),
		crc: view.getUint32(cursor + 16, true),
		compressed: view.getUint32(cursor + 20, true),
		size: view.getUint32(cursor + 24, true),
		local: view.getUint32(cursor + 42, true),
		nameLength,
		next: cursor + 46 + nameLength + view.getUint16(cursor + 30, true) + view.getUint16(cursor + 32, true)
	};
	if (header.next > end || header.flags & 0x2041 ||
		(header.method !== 0 && header.method !== 8) || header.compressed === 0xffffffff ||
		header.size === 0xffffffff || header.local === 0xffffffff || view.getUint16(cursor + 34, true) !== 0)
		invalid();
	return header;
}
function entryLimits(header: EntryHeader, limits: ArchiveLimits, total: number): void {
	if (header.size > limits.entryBytes) throw new XlsxLimitError('A workbook archive entry is too large.');
	if (header.size / Math.max(header.compressed, 1) > limits.compressionRatio)
		throw new XlsxLimitError('The workbook has an unsafe compression ratio.');
	if (total > limits.totalBytes) throw new XlsxLimitError('The uncompressed workbook is too large.');
}
function localEntry(bytes: Uint8Array, view: DataView, header: EntryHeader, name: string, central: number): Entry {
	const { local, flags, method, crc, compressed, size } = header;
	if (local + 30 > central || view.getUint32(local, true) !== 0x04034b50 ||
		view.getUint16(local + 6, true) !== flags || view.getUint16(local + 8, true) !== method)
		invalid();
	const nameLength = view.getUint16(local + 26, true);
	const start = local + 30 + nameLength + view.getUint16(local + 28, true);
	if (start + compressed > central ||
		entryName(bytes.subarray(local + 30, local + 30 + nameLength), flags) !== name)
		invalid();
	if (!(flags & 8) && (view.getUint32(local + 14, true) !== crc ||
		view.getUint32(local + 18, true) !== compressed || view.getUint32(local + 22, true) !== size))
		invalid();
	if (method === 0 && compressed !== size) invalid();
	return { name, method, crc, size, start, end: start + compressed, local };
}
function checkOverlaps(entries: Map<string, Entry>): void {
	let previousEnd = 0;
	for (const entry of [...entries.values()].sort((a, b) => a.local - b.local)) {
		if (entry.local < previousEnd) invalid();
		previousEnd = entry.end;
	}
}
function archiveEntries(bytes: Uint8Array, limits: ArchiveLimits): Map<string, Entry> {
	if (bytes.length > limits.fileBytes)
		throw new XlsxLimitError(`File exceeds the ${limits.fileBytes} bytes import limit.`);
	if (bytes.length < 22) invalid();
	const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
	const { end, central, count } = directory(view, limits);
	const entries = new Map<string, Entry>();
	let cursor = central;
	let total = 0;
	for (let index = 0; index < count; index++) {
		const header = centralHeader(view, cursor, end);
		total += header.size;
		entryLimits(header, limits, total);
		const name = entryName(bytes.subarray(cursor + 46, cursor + 46 + header.nameLength), header.flags);
		if (entries.has(name)) invalid();
		entries.set(name, localEntry(bytes, view, header, name, central));
		cursor = header.next;
	}
	if (cursor !== end) invalid();
	checkOverlaps(entries);
	return entries;
}
/** ZIP32, methods STORE and DEFLATE. All offsets are bounded before access. */
export class ZipArchive {
	private entries: Map<string, Entry>;
	constructor(private bytes: Uint8Array, private limits: ArchiveLimits) {
		this.entries = archiveEntries(bytes, limits);
	}
	async read(
		name: string,
		maximum = this.limits.entryBytes,
		signal?: AbortSignal
	): Promise<Uint8Array> {
		signal?.throwIfAborted();
		const entry = this.entries.get(name);
		if (!entry) throw new XlsxError(`Missing workbook part: ${name}`);
		if (entry.size > maximum) throw new XlsxLimitError('A workbook archive entry is too large.');
		const compressed = this.bytes.subarray(entry.start, entry.end);
		const bytes = entry.method === 8 ? await inflate(compressed, entry.size, signal) : compressed;
		if (bytes.length !== entry.size || crc32(bytes) !== entry.crc)
			throw new XlsxError('Workbook archive checksum or size mismatch.');
		await checkpoint(signal);
		return bytes;
	}
	names(): string[] {
		return [...this.entries.keys()];
	}
}
/** Write DEFLATE-compressed ZIP entries with the platform compression API. */
export async function writeZip(
	parts: ReadonlyMap<string, string | Uint8Array>,
	signal?: AbortSignal
): Promise<Uint8Array> {
	signal?.throwIfAborted();
	if (parts.size >= 65535) throw new XlsxError('ZIP64 workbooks are not supported.');
	const entries: Array<{
		name: Uint8Array;
		data: Uint8Array;
		size: number;
		crc: number;
		offset: number;
	}> = [];
	let localSize = 0,
		centralSize = 0;
	for (const [path, text] of parts) {
		const name = encoder.encode(path),
			input = typeof text === 'string' ? encoder.encode(text) : new Uint8Array(text);
		entryName(name, 0x800);
		if (name.length > 65535) invalid();
		const data = await deflate(input, signal);
		entries.push({ name, data, size: input.length, crc: crc32(input), offset: localSize });
		localSize += 30 + name.length + data.length;
		centralSize += 46 + name.length;
		await checkpoint(signal);
	}
	const total = localSize + centralSize + 22;
	if (total >= 0xffffffff) throw new XlsxError('ZIP64 workbooks are not supported.');
	const bytes = new Uint8Array(total),
		view = new DataView(bytes.buffer);
	let central = localSize;
	for (const entry of entries) {
		const { name, data, size, crc, offset } = entry;
		view.setUint32(offset, 0x04034b50, true);
		view.setUint16(offset + 4, 20, true);
		view.setUint16(offset + 6, 0x800, true);
		view.setUint16(offset + 8, 8, true);
		view.setUint16(offset + 12, 33, true);
		view.setUint32(offset + 14, crc, true);
		view.setUint32(offset + 18, data.length, true);
		view.setUint32(offset + 22, size, true);
		view.setUint16(offset + 26, name.length, true);
		bytes.set(name, offset + 30);
		for (let i = 0; i < data.length; i += 65536) {
			bytes.set(data.subarray(i, i + 65536), offset + 30 + name.length + i);
			await checkpoint(signal);
		}
		view.setUint32(central, 0x02014b50, true);
		view.setUint16(central + 4, 20, true);
		view.setUint16(central + 6, 20, true);
		view.setUint16(central + 8, 0x800, true);
		view.setUint16(central + 10, 8, true);
		view.setUint16(central + 14, 33, true);
		view.setUint32(central + 16, crc, true);
		view.setUint32(central + 20, data.length, true);
		view.setUint32(central + 24, size, true);
		view.setUint16(central + 28, name.length, true);
		view.setUint32(central + 42, offset, true);
		bytes.set(name, central + 46);
		central += 46 + name.length;
	}
	view.setUint32(central, 0x06054b50, true);
	view.setUint16(central + 8, parts.size, true);
	view.setUint16(central + 10, parts.size, true);
	view.setUint32(central + 12, centralSize, true);
	view.setUint32(central + 16, localSize, true);
	return bytes;
}
