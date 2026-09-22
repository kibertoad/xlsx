---
"@office-kit/xlsx": patch
---

A non-finite column width or row height no longer reaches the worksheet part. `setColumnWidth(ws, 1, NaN)` wrote `width="NaN"`, which is no `xsd:double`: the part failed schema validation and Excel offered to repair the file.

- `setColumnWidth` and `setRowHeight` throw `OpenXmlSchemaError` for a size that is not a non-negative finite number.
- Saving throws rather than writing `width="NaN"` / `ht="NaN"` when a non-finite size was put into `ws.columnDimensions` / `ws.rowDimensions` directly. A negative size is left alone, so a file that carries one still round-trips.
- The streaming `setColumnWidth` on a write-only worksheet validated nothing. It now rejects the same sizes, plus a column index outside `[1, 16384]`, and it reports them as `OpenXmlSchemaError` like the modelled setter rather than `OpenXmlIoError`.
- `setColumnWidths` and `setRowHeights` skip a negative entry the way they already skipped a non-finite one, so a bad entry cannot abort the call part-way through.
- `autofitColumn` and `autofitColumns` check `padding` / `min` / `max` before resizing anything, so a `NaN` there is reported against the option the caller passed.

Excel's own ceilings (255 characters, 409 points) are still not enforced: a value past them is a well-formed double that Excel clamps on open, so rejecting one would refuse a file that works.
