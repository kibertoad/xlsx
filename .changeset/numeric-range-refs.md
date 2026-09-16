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
  `moveRange`, `mergeCells`, `unmergeCells`, `makeAutoFilter`
- `@office-kit/xlsx/workbook`: `addDefinedNameForRange`

`writeRange` takes a `{ row, col }` anchor alongside the A1 form. Bounds are validated
against the sheet grid and inverted bounds are normalised, so `{ minRow: 5, maxRow: 1 }`
covers the same cells as `"A5:A1"` instead of iterating nothing. A rejected range is
rejected up front: no cell is written, and no font, fill, border or number format is
registered on the workbook, so a failed `setRangeStyle` no longer leaves records behind
for the next save to serialize. `mergeCells` stores a normalised copy and hands back
another, so neither the argument nor the return value can rewrite a merge that is already
on the sheet.

`getRangeAddress` and `makeAutoFilter` keep the caller's spelling of a range string, `$`
markers and open spans (`"A:E"`) included, since that is what a defined name or a filter
ref wants. Both check it first: a string that is not a range used to travel verbatim into
`workbook.xml` or `sheet1.xml`, and Excel asked to repair the file on open.

`setSelectedRange` keeps its string parameter, because an `sqref` can hold several
ranges, and the `*Str` helpers (`shiftRangeStr`, `rangeAreaStr`, `expandRangeStr` and
friends) stay string-in / string-out by definition.

This widens one parameter rather than adding a second function, so each capability still
has a single canonical helper that reads either spelling, the way `setFreezePanes` now
reads either.

**Breaking:** `setRangeValues` requires `rows` to fit inside `range`. It used to lay
values down from the top-left and run past the bottom and right edges, so
`setRangeValues(ws, 'A1', rows)` wrote a whole block out of a one-cell range; an array
that does not fit now throws instead. `writeRange(ws, 'A1', rows)` is that behaviour
under its own name: it takes an anchor, grows to fit, and returns the box it wrote.

Formula text had the same string-concatenation problem. `tupleToCoordinate` and
`boundariesToRangeString` gained `absoluteCol` / `absoluteRow`, so `$B$5` and
`$A$4:$H$20` come out of the helpers.

**Breaking:** `freezePanes(ws, rows, cols)` is removed. `setFreezePanes` accepts
`'B2' | { rows, cols } | undefined`, which covers both spellings through one function, and
the counts form takes a zero on either axis where `freezePanes` demanded at least 1 on
both. Replace `freezePanes(ws, 1, 1)` with `setFreezePanes(ws, { rows: 1, cols: 1 })`;
`freezeRows` and `freezeColumns` are unchanged. `makeFreezePane` reads the counts too,
and rejects a count past the grid in the caller's terms instead of letting the coordinate
composer complain about a cell the caller never named.
