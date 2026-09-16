---
'@office-kit/xlsx': minor
---

feat: `registerCellStyle` and `patchCellFont`, so styling a report is not a second pass

Two gaps made formatting a generated sheet cost far more calls than it should.

`setCell` already accepted a `styleId` and the stylesheet already deduped xf records,
but there was no way to obtain a `styleId` from a style spec without first having a
cell to hang it on. `registerCellStyle(wb, spec)` returns one:

```ts
const INT = registerCellStyle(wb, { numberFormat: '#,##0', border: THIN });
setCell(ws, row, col, value, INT);
```

The id names a complete style rather than a patch over the target cell: an axis
missing from `spec` renders as the workbook default even where the target cell had
something there. `setCellStyle` and `setRangeStyle` remain the patch-an-existing-cell
paths, so reach for those when the cells already carry formatting you want to keep.

`appendRow` and `appendRows` take matching `{ styleIds }`, positionally aligned with
the values and reused for every row of an `appendRows` call:

```ts
appendRows(ws, rows, { styleIds: [TEXT, INT, INT] });
```

A column with a style id is written even when its value is empty, so a
bordered-but-blank input column survives the append. Ids past a row's last value
therefore add styled blank cells and widen the sheet; trim the array per row
(`styleIds: columnStyles.slice(0, values.length)`) when the input is ragged.

Saving now throws `OpenXmlSchemaError` when a cell's `styleId` names no entry in its
workbook's `cellXfs` pool, naming the sheet and cell. Excel drops such a sheet behind
the repair dialog, and an id reused across two workbooks is the easy way to get there.

`patchCellFont(wb, cell, patch)` merges a partial font over the cell's current one.
`setCellFont(wb, c, makeFont({ bold: true }))` is a legal call that registers a font
with no `<name>` and no `<sz>`, after which Excel, LibreOffice and Sheets each
substitute a different default and nothing warns. A field set to `undefined` is
removed rather than kept, so `setBold`, `setItalic`, `setStrikethrough`,
`setUnderline`, `setFontSize`, `setFontName` and `setFontColor` are all now this
function with one field filled in, rather than seven hand-rolled merges.
