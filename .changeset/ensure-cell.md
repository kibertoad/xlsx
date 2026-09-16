---
'@office-kit/xlsx': minor
---

feat!: `ensureCell` for get-or-create, and `setCell`'s `value` is now mandatory

`setCell(ws, row, col)` read like "reach the cell at (row, col)" and wrote `null`. A
styling pass that walked already-populated rows therefore erased the values and
formulas it touched, and nothing in the signature or the docstring said so. The
cheatsheet and the formula recipe both taught the no-value form as the way to reach a
cell, so the trap was the documented path.

`value` is now required on `setCell` and `setCellByCoord`, which turns the mistake
into a compile error rather than a wrong file. `ensureCell(ws, row, col)` returns the
cell at a coordinate and allocates a blank one only when it does not exist, leaving
an existing value untouched; `ensureCellByCoord(ws, 'B5')` is the A1-addressed form.
That pattern already existed inside `applyToRange`, `setRangeStyle`,
`setRangeWrapText`, `setRangeAlignment` and `setRangeBorderBox`; those now call
`ensureCell` instead of open-coding it.

Migration is mechanical, and the compiler points at every site:

- `setCell(ws, r, c)` becomes `ensureCell(ws, r, c)`, and `setCellByCoord(ws, 'B5')`
  becomes `ensureCellByCoord(ws, 'B5')`
- emptying a cell stays available and is now explicit. `null` is a `CellValue`,
  so `setCell(ws, r, c, null)` clears the value and leaves the cell in the sheet
  with its fill, border and number format, the way Excel's Delete key does.
  `deleteCell` drops the cell outright and `clearRange` does the same across a
  rectangle.

`mergeCells` drops the cells underneath a merge, and reaching one of those
coordinates with `ensureCell` allocates it again, so the written `<sheetData>` carries
a blank `<c>` under the merge. Address the top-left coordinate when the merged block
is what you mean.

Also removed: `setCellFormula`, `setCellArrayFormula` and `setCellRichText` in
`src/worksheet/worksheet.ts`. They were dropped from the public subpaths in an earlier
"one way per task" trim and have been unreachable since. `ensureCell` plus
`setFormula` / `setArrayFormula` covers the two formula wrappers. Rich text is
`setCell(ws, r, c, { kind: 'rich-text', runs: makeRichText(runs) })`: `makeRichText`
builds the runs and the `kind` wrapper is spelled out at the call site, where
`makeErrorValue` / `makeDurationValue` hand back a `CellValue` outright.
