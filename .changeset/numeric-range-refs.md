---
'@office-kit/xlsx': minor
---

feat!: numeric coordinates wherever an A1 string was required, and one way to freeze panes

Every range-taking helper insisted on an A1 string, so code that tracks rows and
columns as integers had to format `"A4:H20"` for the callee to parse straight back
into the numbers it started with. `setRangeStyle`, `setRangeNumberFormat`,
`setRangeBorderBox`, `applyToRange`, `clearRange`, `getRangeValues`,
`setRangeValues`, `copyRange`, `moveRange` and the rest now take
`string | { minRow, minCol, maxRow, maxCol }`, and `writeRange` takes a
`{ row, col }` anchor. The new `RangeRef` type in `@office-kit/xlsx/utils` names the
union. Inverted bounds are normalised, so `{ minRow: 5, maxRow: 1 }` behaves like
`"A5:A1"` instead of iterating nothing.

Formula text has the same problem. `tupleToCoordinate` and
`boundariesToRangeString` gained `absoluteCol` / `absoluteRow`, so `$B$5` and
`$A$4:$H$20` come out of the helpers rather than out of string concatenation.

**Breaking:** `freezePanes(ws, rows, cols)` is removed. `setFreezePanes` now accepts
`'B2' | { rows, cols } | undefined`, which covers both spellings through one
function. The numeric form is also strictly more capable: `freezePanes` required both
counts to be at least 1, so "freeze two rows and no columns" could not be expressed.
Replace `freezePanes(ws, 1, 1)` with `setFreezePanes(ws, { rows: 1, cols: 1 })`.
