import { expect, test } from 'bun:test';
import { Workbook, readWorkbook, writeWorkbook } from '../src/index.js';
import { fixture } from './helpers.js';

for (const operation of ['read', 'write'] as const) {
	for (const reason of [new Error('Cancelled by application'), 'cancelled']) {
		test(`${operation} preserves a pre-aborted signal's ${typeof reason} reason`, async () => {
			const controller = new AbortController();
			controller.abort(reason);
			const book = new Workbook();
			book.addWorksheet('Data');
			const pending = operation === 'read'
				? readWorkbook(await fixture(), {}, controller.signal)
				: writeWorkbook(book, controller.signal);
			await expect(pending).rejects.toBe(reason);
		});
	}

	test(`${operation} yields to application cancellation and preserves its reason`, async () => {
		const controller = new AbortController();
		const reason = new Error('Stop pending workbook work');
		const book = new Workbook();
		book.addWorksheet('Data').addRow(['x'.repeat(100000)]);
		const bytes = await fixture();
		const pending = operation === 'read'
			? readWorkbook(bytes, {}, controller.signal)
			: writeWorkbook(book, controller.signal);
		const timer = setTimeout(() => controller.abort(reason), 0);
		try {
			await expect(pending).rejects.toBe(reason);
		} finally {
			clearTimeout(timer);
		}
	});
}
