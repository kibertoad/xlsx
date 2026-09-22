---
"@office-kit/xlsx": patch
---

Fix an edit to one column dropping the widths of its neighbours. A loaded `<col min="1" max="16384" width="12"/>` is a single entry covering every column, and `setColumnWidth` / `hideColumn` / `groupColumns` deleted the whole run before writing the one column they were asked about. Loading a workbook, changing one column's width and saving lost the width of every other column the run covered. Such a run is now split around the columns being edited, and the rest of it keeps its fields.

The bulk helpers (`hideColumns`, `unhideColumns`, `setColumnWidths`, `groupColumns`, `ungroupColumns`, `collapseColumnGroup`, `expandColumnGroup`) also pair the whole band against the existing runs in one pass instead of scanning the entry map once per column, so their cost grows with the band rather than with its square. `hideColumns(ws, 1, 8000)` drops from about 260 ms to about 3 ms.
