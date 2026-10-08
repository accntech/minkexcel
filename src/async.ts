/** Yield when a cancellation signal needs event-loop access; avoid timers otherwise. */
export function checkpoint(signal?: AbortSignal): Promise<void> {
	signal?.throwIfAborted();
	if (!signal) return Promise.resolve();
	return new Promise((resolve, reject) => {
		const finish = () => {
			signal?.removeEventListener('abort', abort);
			resolve();
		};
		const timer = setTimeout(finish, 0);
		const abort = () => {
			clearTimeout(timer);
			signal?.removeEventListener('abort', abort);
			reject(signal?.reason);
		};
		signal?.addEventListener('abort', abort, { once: true });
	});
}
