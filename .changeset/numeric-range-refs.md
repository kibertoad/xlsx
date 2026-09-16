---
'@office-kit/xlsx': minor
---

feat!: numeric coordinates wherever an A1 string was required, and one way to freeze panes

Every range-taking helper insisted on an A1 string, so code that tracks rows and
columns as integers had to format `"A4:H20"` for the callee to parse straight back
into the numbers it started with. These now take `string | { minRow, minCol, maxRow,
maxCol }`, a union named `RangeRef` in `@office-kit/xlsx/utils`:

- `@office-kit/xlsx/styles`: `setRangeStyle`, `setRangeFont`, `setRangeBackgroundColor`,
  `setRangeNumberFormat`, `setRangeAlignment`, `setRangeWrapText`, `setRangeProtection`,
  `setRangeBorderBox`, `formatAsHeader`, `clearRangeStyle`
- `@office-kit/xlsx/worksheet`: `setRangeValues`, `getRangeValues`, `applyToRange`,
  `clearRange`, `getCellsInRange`, `replaceInRange`, `getRangeAddress`, `copyRange`,
  `moveRange`, `mergeCells`, `unmergeCells`

`writeRange` takes a `{ row, col }` anchor alongside the A1 form. Bounds are validated
against the sheet grid and inverted bounds are normalised, so `{ minRow: 5, maxRow: 1 }`
covers the same cells as `"A5:A1"` instead of iterating nothing, and a fractional or
off-grid bound throws before any cell is touched. `mergeCells` stores a normalised copy,
so mutating a bounds object after the call no longer rewrites a merge that is already on
the sheet.

`setSelectedRange` keeps its string parameter, because an `sqref` can hold several
ranges, and the `*Str` helpers (`shiftRangeStr`, `rangeAreaStr`, `expandRangeStr` and
friends) stay string-in / string-out by definition.

This widens one parameter rather than adding a second function, so each capability still
has a single canonical helper that reads either spelling, the way `setFreezePanes` now
reads either.

**Breaking:** `setRangeValues` clips to the range it was given. Values past the bottom or
right edge of `range` are dropped instead of written outside it, which is what
`copyRange` already does against a smaller target, and it makes `getRangeValues` a true
inverse. `setRangeValues(ws, 'A1', rows)` used to lay down a whole block from a one-cell
range; `writeRange(ws, 'A1', rows)` is that behaviour, and it returns the bounding box it
wrote.

Formula text had the same string-concatenation problem. `tupleToCoordinate` and
`boundariesToRangeString` gained `absoluteCol` / `absoluteRow`, so `$B$5` and
`$A$4:$H$20` come out of the helpers.

**Breaking:** `freezePanes(ws, rows, cols)` is removed. `setFreezePanes` now accepts
`'B2' | { rows, cols } | undefined`, which covers both spellings through one function.
The numeric form is also strictly more capable: `freezePanes` required both counts to be
at least 1, so "freeze two rows and no columns" could not be expressed. Replace
`freezePanes(ws, 1, 1)` with `setFreezePanes(ws, { rows: 1, cols: 1 })`.
