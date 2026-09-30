---
"@office-kit/xlsx": patch
---

`mergeCells` and `clearRange` no longer cost time proportional to the area of the range: the work is bounded by the populated cells on the sheet. Both walked every coordinate of the rectangle, so merging a band Excel lets a caller name was unusable at scale: `mergeCells(ws, 'A:J')` took 57 ms on a sheet holding 20 cells, and a whole-sheet `mergeCells(ws, 'A1:XFD1048576')` walked seventeen billion coordinates and effectively never returned. Both now finish in well under a millisecond on the same sheet.

Each axis is enumerated whichever way is smaller, so a small range on a sheet with hundreds of thousands of rows stays cheap too.
