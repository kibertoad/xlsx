---
'@office-kit/xlsx': minor
---

feat: read a cell as the text Excel shows, and a date-formatted cell as a `Date`

`getCellDisplayText(wb, cell)` from `@office-kit/xlsx/styles` puts a cell's
value through the number format its style points at, so `0.5` under `0.0%`
reads `50.0%`, `1234567.891` under `#,##0.00` reads `1,234,567.89`, and
`1.1 + 2.2` under `General` reads `3.3` instead of `3.3000000000000003`. It
covers the built-in format catalogue plus the common custom codes: digit
placeholders, thousands grouping, percent, currency, scientific and
engineering notation, fractions, the positive / negative / zero / text sections
of a multi-section code, and date, time and elapsed-time codes. A code outside
that set falls back to `cellValueAsString` rather than printing a guess; the
docstring lists the boundary.

`getCellDate(wb, cell)` reads a date-formatted cell as a `Date`. Excel stores a
date as a plain day count, so the number format is the only evidence that
`45365` means 2024-03-14; this resolves the format, checks that it names a
calendar date rather than a time of day (`h:mm`) or an elapsed span
(`[h]:mm:ss`), and converts under the workbook epoch.

New guide at `docs/migrate-from-sheetjs.md`, and a recipe for reading a
workbook somebody else produced.
