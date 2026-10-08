import { expect, test } from 'bun:test';
import { inflateRawSync } from 'node:zlib';
import { deflate } from '../src/deflate.js';

for (const length of [0, 1, 65535, 65536, 65537, 200000]) {
	test(`produces independently decodable raw DEFLATE for ${length} bytes`, async () => {
		let seed = 12345;
		const input = Uint8Array.from({ length }, () => {
			seed = (Math.imul(seed, 1664525) + 1013904223) >>> 0;
			return seed >>> 24;
		});
		const compressed = await deflate(input);
		expect(new Uint8Array(inflateRawSync(compressed))).toEqual(input);
	});
}
test('cancels while feeding the compressor and preserves the abort reason', async () => {
	const controller = new AbortController(),
		reason = new Error('Stop compression');
	const pending = deflate(new Uint8Array(5_000_000), controller.signal);
	setTimeout(() => setTimeout(() => controller.abort(reason), 0), 0);
	await expect(pending).rejects.toBe(reason);
});
