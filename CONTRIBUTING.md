# Contributing to MinkExcel

Please read and follow the [Code of Conduct](CODE_OF_CONDUCT.md) when
participating in the project. It describes community expectations and how to
report concerns.

The library manifest has no runtime, development or peer dependencies.
TypeScript, Playwright and the ExcelJS benchmark reference are isolated in
`tools/package.json`, with their own lockfile. Unit tests use Bun's built-in
runner and Node's built-in compressor/inflater as independent DEFLATE references.

## Commit conventions

Follow the repository's existing commit message format:

```text
type: short description
```

Use a lowercase type and a concise, imperative description, such as
`perf: enable tree-shaking and verify consumer bundles`. Existing types include
`feat`, `perf`, `chore`, `bench`, `ci` and `design`; use `fix` for bug fixes,
`docs` for documentation and `test` for test changes.

Keep each commit focused on one logical change. Separate unrelated changes into
their own commits, and use the commit body when the reason or tradeoffs need
more explanation.

### Verified commits

All commits submitted in a pull request must be signed and show **Verified** on
GitHub. Configure a GPG, SSH or S/MIME signing key using
[GitHub's commit signature verification guide](https://docs.github.com/en/authentication/managing-commit-signature-verification/about-commit-signature-verification).
For GPG or SSH, add the public signing key to your GitHub account and configure
Git to use it.

After configuring your signing key, enable signing for this repository or sign
each commit explicitly:

```sh
git config commit.gpgsign true
git commit -S -m "docs: document contribution requirements"
```

After pushing your branch, check that every commit in the pull request's
**Commits** tab has the **Verified** badge. Recheck after amending or rebasing
commits, since those operations create new commits that also need signatures.
A `Signed-off-by` trailer (`git commit -s`) does not provide a cryptographic
signature and does not satisfy this requirement. See
[GitHub's signing instructions](https://docs.github.com/en/authentication/managing-commit-signature-verification/signing-commits)
for setup and signing details.

## Build and unit tests

Run from the repository root. The build needs Node.js; unit tests and the
browser/benchmark scripts need Bun. The current release workflow uses Node.js 24
and Bun 1.4.0.

```sh
bun install --cwd tools --frozen-lockfile --ignore-scripts
bun run build
bun run test
```

The build writes ES modules and TypeScript declarations to `dist/`.
Unit tests and synthetic workbook fixtures live in `tests/`; unit tests can also
run without installing the optional tools.

Run `bun run test:coverage` to inspect unit-test coverage. Tests are grouped by
the public worksheet model, import, export, cancellation, XML and ZIP/DEFLATE
handling. Worksheet import cases exercise both the visitor-based path without
a signal and the cancellable path with a signal. Independent ExcelJS fixtures,
a stored ZIP fixture and Node's zlib help catch reader/writer defects that a
roundtrip through MinkExcel alone could hide.

## Browser and Web Worker checks

Install Chromium before running the harness:

```sh
bun tools/node_modules/@playwright/test/cli.js install chromium
bun run test:browser
```

`tests/browser.ts` builds through a consumer entry and tests compressed import/export,
a Blob roundtrip, scalar values and a module Web Worker. It uses a local server
and Playwright's Chromium. This harness does not verify Firefox or Safari.

## ExcelJS interoperability and integrity

```sh
bun run build
bun run test:interop
```

This checks the built package against the pinned ExcelJS version in `tools/`.
Both readers verify text preservation from both writers. Identical malformed
files exercise CRC32, truncated ZIP, reserved XML namespace and DTD handling.
The command asserts MinkExcel's expected behavior and records the comparison in
`benchmarks/RELIABILITY.md` and `benchmarks/reliability.json`.

Push, pull request and release workflows run build, unit, tree-shaking,
interoperability, browser and packaging checks. Timing benchmarks run separately
because hardware and runtime load affect results.

## Benchmarks

Run on an otherwise idle machine after building. Run commands sequentially;
concurrent benchmarks distort timings. Set `XLSX_BENCH_ITERATIONS` to change the
sample count (default seven for the workload matrix). The recorded comparisons
use eleven samples after one warmup:

```sh
bun run build
XLSX_BENCH_ITERATIONS=11 bun run bench:matrix
XLSX_BENCH_ITERATIONS=11 npm run bench:node
```

`benchmarks/matrix.ts` tests numeric, text and mixed data at 100, 1,000 and
10,000 rows with eight columns. It imports the built ES modules, includes
workbook construction in export timing and feeds identical compressed input to
both readers. Every data cell is verified with both readers outside timings,
including a MinkExcel roundtrip. Node requires native TypeScript type stripping;
the command was verified with Node 24.12.0.

Bun writes `benchmarks/MATRIX.md` and `benchmarks/matrix.json`; Node writes
`benchmarks/MATRIX-NODE.md` and `benchmarks/matrix-node.json`. JSON includes raw
timing samples and output sizes. Review runtime, hardware, warmups, sample count
and scope differences with each result. No AbortSignal is supplied. MinkExcel
imports values; ExcelJS also builds a richer presentation model. These checks
do not measure streaming, peak memory or browser speed.

The original mixed report fixture remains available for comparison with the
pre-optimization measurements in `benchmarks/baseline.json`:

```sh
XLSX_BENCH_ITERATIONS=11 bun run bench
```

It updates `benchmarks/RESULTS.md` and `benchmarks/results.json`. Keep the
historical baseline unchanged when recording a new result.

## Browser bundle size

Run this from the repository root to check tree-shaking and measure consumer
bundles with Bun's browser ESM bundler and minifier. Gzip uses Node's built-in
compressor defaults:

```sh
bun run build
bun run test:treeshaking
```

`tests/treeshaking.ts` copies the built package and its manifest into a temporary
consumer's `node_modules/`. It checks model-only, export-only, import-only,
unused-import and full bundles through the public `minkexcel` entry point,
asserts that unused code is absent and roundtrips values between the separate
model, writer and reader bundles. The temporary files are removed afterward.

Keep module initialization free of externally visible side effects to preserve
the package's `sideEffects: false` declaration. Pure annotations belong only on
local initialization that can safely be omitted when its result is unused.

Record the bundler version and measurement date alongside size claims.
This measures JavaScript, excluding TypeScript declarations and the icon.

## Format references and publishing

ZIP layout follows the
[PKWARE APPNOTE](https://pkware.cachefly.net/webdocs/casestudies/APPNOTE.TXT);
DEFLATE follows [RFC 1951](https://www.rfc-editor.org/rfc/rfc1951).
Export compression uses the
[Compression Standard](https://compression.spec.whatwg.org/#supported-formats).

See [RELEASING.md](RELEASING.md) for package previews and publishing.
