import { cp, mkdtemp, readFile, readdir, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join, resolve } from 'node:path';
import { createInstrumenter } from '../node_modules/istanbul-lib-instrument';

const root = resolve(import.meta.dir, '../..');
const output = join(root, 'build/coverage/coverage-final.json');
const directory = await mkdtemp(join(tmpdir(), 'minkexcel-coverage-'));
try {
	await rm(output, { force: true });
	await cp(join(root, 'src'), join(directory, 'src'), { recursive: true });
	await cp(join(root, 'tests'), join(directory, 'tests'), { recursive: true });
	for (const name of await readdir(join(root, 'src'))) {
		if (!name.endsWith('.ts')) continue;
		const path = join(root, 'src', name);
		const instrumenter = createInstrumenter({
			esModules: true,
			parserPlugins: ['typescript'],
			coverageGlobalScope: 'process',
			coverageGlobalScopeFunc: false,
			coverageVariable: '__minkexcelCoverage__'
		});
		const source = instrumenter.instrumentSync(await readFile(path, 'utf8'), path);
		await writeFile(join(directory, 'src', name), source);
	}
	const child = Bun.spawn(['bun', 'test', '--preload', join(import.meta.dir, 'preload.ts')], {
		cwd: directory,
		env: { ...process.env, MINKEXCEL_COVERAGE: output },
		stdout: 'inherit', stderr: 'inherit'
	});
	process.exitCode = await child.exited;
} finally {
	await rm(directory, { recursive: true, force: true });
}
