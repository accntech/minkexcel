# XLSX benchmark against ExcelJS

2026-10-08T10:27:35.515Z; Bun 1.4.0; darwin arm64; Apple M4 Pro; ExcelJS 4.4.0.

Median of 11 samples after one warmup for each case. Synthetic general ledger: eight columns, strings with leading zeros, dates, amounts, cached formulas, widths, frozen header, number formats and print settings. Export includes workbook construction and serialization. Import uses the same compressed ExcelJS output for both readers. Output checks run outside timings. MinkExcel reads values; ExcelJS also builds its richer presentation model, so these are different feature sets.

| Data rows | MinkExcel export ms | ExcelJS export ms | MinkExcel import ms | ExcelJS import ms | MinkExcel KiB | ExcelJS KiB |
| ---: | ---: | ---: | ---: | ---: | ---: | ---: |
| 100 | 1.2 | 3.6 | 1.9 | 2.6 | 8.8 | 11.7 |
| 1,000 | 9.7 | 15.3 | 11.4 | 14.0 | 52.1 | 56.1 |
| 10,000 | 90.7 | 150.4 | 92.3 | 126.1 | 485.3 | 490.7 |

Both writers use DEFLATE. MinkExcel uses the platform CompressionStream API; ExcelJS uses its bundled ZIP compressor. Compression settings and workbook XML differ. These are local Bun measurements, not browser measurements or performance guarantees.

Run from the repository root: `bun run bench`. Install optional references with `bun install --cwd tools` first. Set `XLSX_BENCH_ITERATIONS` to change the sample count.
