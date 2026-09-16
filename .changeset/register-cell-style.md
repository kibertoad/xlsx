---
'@office-kit/xlsx': minor
---

feat: `registerCellStyle` and `patchCellFont`, so styling a report is not a second pass

Two gaps made formatting a generated sheet cost far more calls than it should.

`setCell` already accepted a `styleId` and the stylesheet already deduped xf records,
but there was no way to obtain a `styleId` from a style spec without first having a
cell to hang it on. `setRangeStyle` built exactly that xf internally and threw the id
away. `registerCellStyle(wb, spec)` now returns it, so a report can build each look
once and let the write carry the formatting:

```ts
const INT = registerCellStyle(wb, { numberFormat: '#,##0', border: THIN });
setCell(ws, row, col, value, INT);
```

`appendRow` and `appendRows` take a matching `{ styleIds }` option, positionally
aligned with the values. A column with a style id is written even when its value is
empty, so a bordered-but-blank input column survives the append.

The id names a complete style rather than a patch over the target cell;
`setCellStyle` remains the patch-an-existing-cell path.

`patchCellFont(wb, cell, patch)` merges a partial font over the cell's current one.
`setCellFont(wb, c, makeFont({ bold: true }))` is a legal call that registers a font
with no `<name>` and no `<sz>`, after which Excel, LibreOffice and Sheets each
substitute a different default and nothing warns. `setBold`, `setItalic`,
`setStrikethrough`, `setFontSize`, `setFontName` and `setFontColor` are now this
function with one field filled in, so there is a single merge path rather than six.
