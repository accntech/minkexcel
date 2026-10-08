import { strict as assert } from 'node:assert';
import { cp, mkdtemp, mkdir, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { gzipSync } from 'node:zlib';
import type { Workbook } from '../src/index.js';

// Exercise the published ESM layout and package metadata through real consumer imports.
const directory = await mkdtemp(join(tmpdir(), 'minkexcel-treeshaking-'));
const packageDirectory = join(directory, 'node_modules', 'minkexcel');
const cases = [
	{
		name: 'model',
		entry: 'export { Workbook } from "minkexcel";',
		absent: ['readWorkbook', 'writeWorkbook', 'Huffman', 'TextEncoder', 'TextDecoder', 'Uint32Array.from', 'CompressionStream', 'Invalid workbook XML.']
	},
	{
		name: 'write',
		entry: 'export { Workbook, writeWorkbook } from "minkexcel";',
		absent: ['readWorkbook', 'ZipArchive', 'Huffman', 'Invalid workbook XML.']
	},
	{
		name: 'read',
		entry: 'export { readWorkbook } from "minkexcel";',
		absent: ['writeWorkbook', 'CompressionStream', 'TextEncoder']
	},
	{
		name: 'unused',
		entry: 'import { readWorkbook } from "minkexcel"; export const sentinel = 42;',
		absent: ['readWorkbook', 'Huffman', 'TextEncoder', 'TextDecoder', 'Uint32Array.from', 'new Map']
	},
	{ name: 'full', entry: 'export * from "minkexcel";', absent: [] }
];

type Consumer = {
	Workbook?: new () => Workbook;
	readWorkbook?: (bytes: Uint8Array) => Promise<Workbook>;
	writeWorkbook?: (book: Workbook) => Promise<Uint8Array>;
	sentinel?: number;
};
const consumers = new Map<string, Consumer>();

try {
	await mkdir(packageDirectory, { recursive: true });
	await cp(new URL('../dist/', import.meta.url), join(packageDirectory, 'dist'), { recursive: true });
	await cp(new URL('../package.json', import.meta.url), join(packageDirectory, 'package.json'));
	console.log(`Browser ESM bundles, Bun ${Bun.version} (minified / gzip bytes):`);
	for (const { name, entry, absent } of cases) {
		const entrypoint = join(directory, `${name}.js`);
		await writeFile(entrypoint, entry);
		const options = { entrypoints: [entrypoint], target: 'browser' as const, format: 'esm' as const };
		const build = await Bun.build(options);
		if (!build.success) throw new AggregateError(build.logs, `${name} bundle failed`);
		const code = await build.outputs[0].text();
		for (const marker of absent)
			assert.ok(!code.includes(marker), `${name} bundle retained unused code: ${marker}`);
		const minified = await Bun.build({ ...options, minify: true });
		if (!minified.success) throw new AggregateError(minified.logs, `${name} minification failed`);
		const bytes = new Uint8Array(await minified.outputs[0].arrayBuffer());
		consumers.set(name, await import(`data:text/javascript;base64,${Buffer.from(bytes).toString('base64')}`));
		console.log(`${name}: ${bytes.length} / ${gzipSync(bytes).length}`);
	}

	assert.equal(consumers.get('unused')!.sentinel, 42);
	const book = new (consumers.get('model')!.Workbook!)();
	book.addWorksheet('Data').addRow(['001234567', 'α & <Co> 🧾', 12.34]);
	const bytes = await consumers.get('write')!.writeWorkbook!(book);
	const imported = await consumers.get('read')!.readWorkbook!(bytes);
	assert.deepEqual(imported.worksheets[0].getRow(1).values.slice(1), ['001234567', 'α & <Co> 🧾', 12.34]);
	const full = consumers.get('full')!;
	const roundtrip = await full.readWorkbook!(await full.writeWorkbook!(imported));
	assert.deepEqual(roundtrip.worksheets[0].getRow(1).values, imported.worksheets[0].getRow(1).values);
	console.log('Unused code excluded; separate model, writer and reader bundles roundtrip correctly.');
} finally {
	await rm(directory, { recursive: true, force: true });
}
