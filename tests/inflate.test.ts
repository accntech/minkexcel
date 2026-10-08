import { describe, expect, test } from 'bun:test';
import { deflateRawSync, constants } from 'node:zlib';
import { inflate } from '../src/inflate.js';

// Independent platform compressor exercises our reader; no external test packages.
describe('raw DEFLATE reader', () => {
	for (const options of [
		{ level: 0 },
		{ strategy: constants.Z_FIXED },
		{ level: 6 },
		{ level: 9 }
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
	test('rejects output beyond the bound and truncated or invalid streams', async () => {
		await expect(
			inflate(deflateRawSync(new TextEncoder().encode('a'.repeat(10000))), 20)
		).rejects.toThrow('large');
		for (const data of [new Uint8Array(), new Uint8Array([7]), new Uint8Array([1, 2, 0, 0, 0])])
			await expect(inflate(data, 100)).rejects.toThrow();
	});
	test('yields during large work so cancellation stops decompression', async () => {
		const controller = new AbortController();
		const pending = inflate(
			deflateRawSync(new Uint8Array(2_000_000)),
			2_000_000,
			controller.signal
		);
		setTimeout(() => controller.abort(), 0);
		await expect(pending).rejects.toThrow();
	});
});
