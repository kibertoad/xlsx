---
"@office-kit/xlsx": patch
---

A non-finite column width or row height no longer reaches the worksheet part. Excel opens a part carrying `width="NaN"` without complaint and turns the column into `width="0" hidden="1"`, so `setColumnWidth(ws, 1, NaN)` produced a workbook whose column had silently disappeared; `ht="NaN"` fell back to the default row height the same way.

- `setColumnWidth` and `setRowHeight` throw `OpenXmlSchemaError` for a size that is not a non-negative finite number.
- Saving throws rather than writing `width="NaN"` / `ht="NaN"` when a non-finite size was put into `ws.columnDimensions` / `ws.rowDimensions` directly. A negative size is left alone, so a file that carries one still round-trips.
- The streaming `setColumnWidth` on a write-only worksheet validated nothing. It now rejects the same sizes, plus a column index outside `[1, 16384]`, and reports them as `OpenXmlSchemaError` like the modelled setter.
- `setColumnWidths` and `setRowHeights` skip a negative entry the way they already skipped a non-finite one, so a bad entry cannot abort the call part-way through.
- `autofitColumn` and `autofitColumns` check `padding` / `min` / `max` before resizing anything, so a `NaN` there is reported against the option the caller passed.

Excel's own ceilings (255 characters, 409 points) are still not enforced: Excel keeps a value past them as written rather than refusing it, so rejecting one would refuse a file that opens fine.
