import { afterAll } from 'bun:test';
import { mkdirSync, writeFileSync } from 'node:fs';
import { dirname } from 'node:path';

const output = process.env.MINKEXCEL_COVERAGE;
if (!output) throw new Error('Run coverage through tools/instrumentation/run.ts.');
afterAll(() => {
	const coverage = (process as typeof process & {
		__minkexcelCoverage__?: Record<string, unknown>;
	}).__minkexcelCoverage__;
	if (!coverage || !Object.keys(coverage).length) throw new Error('No source coverage was collected.');
	mkdirSync(dirname(output), { recursive: true });
	writeFileSync(output, JSON.stringify(coverage) + '\n');
});
