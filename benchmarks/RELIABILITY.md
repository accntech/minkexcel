# File integrity and string preservation comparison

2026-10-08T10:28:32.267Z; Bun 1.4.0; ExcelJS 4.4.0.

String checks compare each writer's output with both independent readers against the original values. Integrity checks feed identical bytes to both readers with default public import options. The CRC case changes the declared checksum in both ZIP headers, leaving payload bytes unchanged. It tests whether the reader validates the declared checksum against actual data. The DTD case contains an internal declaration; no external resources are used.

| Check | MinkExcel | ExcelJS 4.4.0 |
| --- | --- | --- |
| Preserve string "_x0041_" with both readers | Preserved | Changed |
| Preserve string "_x005F_" with both readers | Preserved | Changed |
| Preserve string "first\rsecond" with both readers | Preserved | Changed |
| Preserve string "  α & <Co> 🧾  " with both readers | Preserved | Preserved |
| Reject mismatched worksheet CRC32 (both ZIP headers changed consistently) | Rejected | Accepted |
| Reject truncated ZIP archive | Rejected | Rejected |
| Reject reserved XML namespace rebinding | Rejected | Accepted |
| Reject worksheet DTD | Rejected | Accepted |

These are specific reproducible checks, not an overall reliability score. ExcelJS supports a broader workbook model; passing these checks does not establish feature parity or reliability for every workbook. Reproduce with `bun run test:interop` after installing tools and building the package.
