---
'@office-kit/xlsx': patch
---

fix: `setCellAtAddress` / `getCellAtAddress` say which argument is wrong

Passing a Worksheet where the sheet-qualified address goes (reaching for a
`setCellAtAddress(wb, ws, 'A1', value)` signature that does not exist) used to
fail with `parseSheetRange: missing "!" delimiter in "[object Object]"`, naming
a helper the caller never called and a delimiter that was not the problem. Both
functions now report that the address has to be a string such as `"Sheet1!A1"`,
what they received instead, and point at `getCellByCoord` / `setCellByCoord`
for the case where the worksheet is already in hand. `parseSheetRange` names
the type it was handed rather than stringifying it.
