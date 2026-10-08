/** Yield to the browser/server event loop and release the timer on cancellation. */
export function checkpoint(signal?: AbortSignal): Promise<void> {
	signal?.throwIfAborted();
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
