---
"@office-kit/xlsx": minor
---

feat: `getValueExtent` and `iterRows({ extent: 'values' })`, so a sheet formatted past its data no longer yields rows of `null`

A cell that exists only to carry a style counts towards `getDataExtent`,
`getMaxRow` and `getMaxCol`, which is the right answer for Excel's used range
and the `<dimension>` element, but means a sheet with formatting on 200 rows
and values in 4 iterates 200 rows, 196 of them entirely `null`.
`getValueExtent` reports the bounding box of the cells that hold a value, and
`iterRows` / `iterValues` accept `extent: 'values'` to be bounded by it. The
default is unchanged.

Also newly exported: `getSheetByIndex` from `@office-kit/xlsx/workbook`, so
reading the first sheet of an uploaded file no longer needs a `sheetNames`
lookup and two non-null assertions, and `getDataExtentRef` from
`@office-kit/xlsx/worksheet`, which is the `ref` string `makeAutoFilter` and
`makeTableDefinition` want.
