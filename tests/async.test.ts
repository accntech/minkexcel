import { expect, test } from 'bun:test';
import { checkpoint } from '../src/async.js';

test('operations without a cancellation signal do not wait for a timer at each checkpoint', async () => {
	let timerRan = false;
	const timer = setTimeout(() => { timerRan = true; }, 0);
	try {
		await checkpoint();
		expect(timerRan).toBe(false);
	} finally {
		clearTimeout(timer);
	}
});

test('signal checkpoints yield so an application can cancel and retain its reason', async () => {
	const controller = new AbortController();
	const reason = new Error('Cancelled by user');
	const timer = setTimeout(() => controller.abort(reason), 0);
	try {
		await expect(checkpoint(controller.signal)).rejects.toBe(reason);
	} finally {
		clearTimeout(timer);
	}
});
