---
"@office-kit/xlsx": minor
---

`setPrintArea` now writes a sheet-qualified value. It previously stored the `ref` verbatim, so `setPrintArea(wb, 0, 'A1:E20')` produced `_xlnm.Print_Area` = `A1:E20`, which Excel resolves against whatever sheet is active rather than the one the name is scoped to. The sheet title is taken from `sheetIndex`, so `'A1:E20'` becomes `Report!A1:E20` and each leg of a multi-area range is qualified on its own. An already-qualified `ref` is stored as given.

`setPrintArea` also throws `OpenXmlSchemaError` when `sheetIndex` names no sheet on the workbook, instead of registering a name scoped past the end of the sheet list.

Callers reading `_xlnm.Print_Area` values back out of `listPrintAreas` or `getDefinedName` will see the qualified form. `getDefinedNameTarget` can now resolve those values, which it could not before.
