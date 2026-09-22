---
"@office-kit/xlsx": patch
---

`setColumnWidth`, `setRowHeight`, `setColumnDimension` and `setRowDimension` now reject a non-finite or negative size instead of writing it into the part. A `NaN` width reached the worksheet as `width="NaN"`, which is not a valid `xsd:double`: the part failed schema validation and Excel offered to repair the file. Excel's own ceilings (255 characters, 409 points) are still not enforced, since a value past them is a well-formed double that Excel clamps on open.

The streaming `setColumnWidth` on a write-only worksheet validated nothing at all. It now rejects a non-finite width and an off-grid column index, which previously reached the part as `<col min="0" max="0" width="NaN"/>`.
