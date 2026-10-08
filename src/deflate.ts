import { checkpoint } from './async.js';

/** CompressionStream's zlib envelope has a two-byte header and four-byte Adler32 trailer.
 * Its dictionary flag is unsupported, so removing that envelope yields raw ZIP DEFLATE.
 */
export async function deflate(
	input: Uint8Array<ArrayBuffer>,
	signal?: AbortSignal
): Promise<Uint8Array> {
	signal?.throwIfAborted();
	let offset = 0;
	const source = new ReadableStream<Uint8Array<ArrayBuffer>>({
		async pull(controller) {
			await checkpoint(signal);
			if (offset === input.length) {
				controller.close();
				return;
			}
			controller.enqueue(input.subarray(offset, offset + 65536));
			offset = Math.min(offset + 65536, input.length);
		}
	});
	const reader = source.pipeThrough(new CompressionStream('deflate'), { signal }).getReader();
	const chunks: Uint8Array[] = [];
	let size = 0;
	try {
		for (;;) {
			const { value, done } = await reader.read();
			signal?.throwIfAborted();
			if (done) break;
			chunks.push(value);
			size += value.length;
		}
		const output = new Uint8Array(size);
		let cursor = 0;
		for (const chunk of chunks) {
			output.set(chunk, cursor);
			cursor += chunk.length;
		}
		return output.subarray(2, size - 4);
	} finally {
		// If read failed, cancellation may reject with the same stream error; preserve the original failure.
		await reader.cancel().catch(() => {});
		reader.releaseLock();
	}
}
