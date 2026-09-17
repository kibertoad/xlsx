---
"@office-kit/xlsx": patch
---

`loadWorkbookStream` now finds `sharedStrings` and `styles` wherever the
workbook relationships point, matching `loadWorkbook`. It previously looked
only at `xl/sharedStrings.xml` and `xl/styles.xml`, so a workbook that stores
either part elsewhere (legal, and what some producers emit) streamed back
`null` for every shared-string cell and resolved every style against an empty
pool, with no error to signal it. Both loaders share one lookup now.
