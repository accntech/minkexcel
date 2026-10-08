# ExcelJS comparison across workloads

2026-10-08T10:26:31.101Z; Bun 1.4.0; darwin arm64; Apple M4 Pro; ExcelJS 4.4.0.

Median of 11 samples after one warmup. Uses the built package ES modules. Each worksheet has eight columns and a header. Numeric data includes fractions, negatives and zeros; text includes leading-zero identifiers, Unicode, XML characters, literal formula-like text and whitespace; mixed data includes strings, dates, booleans, numbers, cached formulas, blanks and errors. Both writers apply the same supported styles. Export includes workbook construction. Both readers receive the same compressed ExcelJS output. No AbortSignal is supplied. Every data cell is checked with both readers outside timed samples, including a MinkExcel roundtrip. Raw samples and file sizes are in matrix.json. Literal Excel escape sequences are tested separately by the reliability comparison because ExcelJS 4.4.0 does not preserve them in these checks.

| Workload | Data rows | MinkExcel export ms | ExcelJS export ms | MinkExcel import ms | ExcelJS import ms |
| --- | ---: | ---: | ---: | ---: | ---: |
| numeric | 100 | 0.8 | 3.3 | 1.6 | 2.3 |
| numeric | 1,000 | 5.9 | 12.4 | 8.4 | 11.5 |
| numeric | 10,000 | 56.8 | 114.5 | 68.5 | 96.3 |
| text | 100 | 0.8 | 2.5 | 1.4 | 1.6 |
| text | 1,000 | 6.7 | 13.2 | 11.2 | 12.4 |
| text | 10,000 | 70.1 | 130.9 | 85.5 | 121.7 |
| mixed | 100 | 0.8 | 2.1 | 1.0 | 1.6 |
| mixed | 1,000 | 6.8 | 13.5 | 7.9 | 11.5 |
| mixed | 10,000 | 69.9 | 133.7 | 79.9 | 118.1 |

MinkExcel reads values while ExcelJS builds a richer presentation model. These local document-model results do not compare ExcelJS streaming, browser performance or peak memory, and do not establish a universal speed advantage. Reproduce with `bun run bench:matrix` after installing tools and building the package. The Node command requires a runtime with native TypeScript type stripping; it was verified on Node 24.12.0.
