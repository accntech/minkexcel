import { describe, expect, test } from 'bun:test';
import { deflateRawSync, constants } from 'node:zlib';
import { inflate } from '../src/inflate.js';
import { XlsxError, XlsxLimitError } from '../src/index.js';

// Independent platform compressor exercises our reader; no external test packages.
describe('raw DEFLATE reader', () => {
	for (const options of [
		{ level: 0 },
		{ strategy: constants.Z_FIXED },
		{ level: 6 }
	]) {
		test(`decodes stored, fixed and dynamic blocks ${JSON.stringify(options)}`, async () => {
			for (const text of [
				'',
				'hello',
				'α & <Co> 🧾'.repeat(9000),
				Array.from({ length: 70000 }, (_, i) => String.fromCharCode((i * 31 + i * i) % 256)).join(
					''
				)
			]) {
				const input = new TextEncoder().encode(text);
				expect(await inflate(deflateRawSync(input, options), input.length)).toEqual(input);
			}
		});
	}
	test('accepts an exact output bound and rejects expansion one byte beyond it for every block type', async () => {
		const input = new TextEncoder().encode('a'.repeat(10000));
		for (const options of [{ level: 0 }, { strategy: constants.Z_FIXED }, { level: 6 }]) {
			const compressed = deflateRawSync(input, options);
			expect(await inflate(compressed, input.length)).toEqual(input);
			await expect(inflate(compressed, input.length - 1)).rejects.toThrow(XlsxLimitError);
		}
	});

	test('rejects truncated or invalid streams and bytes trailing a valid stream', async () => {
		for (const data of [new Uint8Array(), new Uint8Array([7]), new Uint8Array([1, 2, 0, 0, 0])])
			await expect(inflate(data, 100)).rejects.toThrow(XlsxError);
		const compressed = deflateRawSync(new TextEncoder().encode('hello'));
		await expect(inflate(compressed.subarray(0, compressed.length - 1), 100)).rejects.toThrow(XlsxError);
		await expect(inflate(new Uint8Array([...compressed, 0]), 100)).rejects.toThrow(XlsxError);
	});
	test('yields during large work so cancellation stops decompression', async () => {
		const controller = new AbortController();
		const reason = new Error('Stop decompression');
		const pending = inflate(
			deflateRawSync(new Uint8Array(8_000_000)),
			8_000_000,
			controller.signal
		);
		setTimeout(() => setTimeout(() => controller.abort(reason), 0), 0);
		await expect(pending).rejects.toBe(reason);
	});
});
