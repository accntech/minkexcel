# ExcelJS comparison across workloads

2026-10-08T10:25:31.858Z; Node 24.12.0; darwin arm64; Apple M4 Pro; ExcelJS 4.4.0.

Median of 11 samples after one warmup. Uses the built package ES modules. Each worksheet has eight columns and a header. Numeric data includes fractions, negatives and zeros; text includes leading-zero identifiers, Unicode, XML characters, literal formula-like text and whitespace; mixed data includes strings, dates, booleans, numbers, cached formulas, blanks and errors. Both writers apply the same supported styles. Export includes workbook construction. Both readers receive the same compressed ExcelJS output. No AbortSignal is supplied. Every data cell is checked with both readers outside timed samples, including a MinkExcel roundtrip. Raw samples and file sizes are in matrix-node.json. Literal Excel escape sequences are tested separately by the reliability comparison because ExcelJS 4.4.0 does not preserve them in these checks.

| Workload | Data rows | MinkExcel export ms | ExcelJS export ms | MinkExcel import ms | ExcelJS import ms |
| --- | ---: | ---: | ---: | ---: | ---: |
| numeric | 100 | 1.7 | 3.0 | 1.7 | 2.5 |
| numeric | 1,000 | 7.8 | 15.9 | 6.9 | 9.9 |
| numeric | 10,000 | 60.7 | 167.7 | 64.1 | 85.5 |
| text | 100 | 1.3 | 2.9 | 1.5 | 1.9 |
| text | 1,000 | 6.6 | 17.9 | 8.4 | 10.2 |
| text | 10,000 | 66.6 | 186.1 | 80.9 | 94.2 |
| mixed | 100 | 1.5 | 3.5 | 0.9 | 1.8 |
| mixed | 1,000 | 6.8 | 18.1 | 7.0 | 9.5 |
| mixed | 10,000 | 69.3 | 192.2 | 75.5 | 93.4 |

MinkExcel reads values while ExcelJS builds a richer presentation model. These local document-model results do not compare ExcelJS streaming, browser performance or peak memory, and do not establish a universal speed advantage. Reproduce with `npm run bench:node` after installing tools and building the package. The Node command requires a runtime with native TypeScript type stripping; it was verified on Node 24.12.0.
