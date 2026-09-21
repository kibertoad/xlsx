---
'@office-kit/xlsx': patch
---

fix: `setCellAtAddress` / `getCellAtAddress` say which argument is wrong

Reaching for a `setCellAtAddress(wb, ws, 'A1', value)` signature that does not
exist used to fail with `parseSheetRange: missing "!" delimiter in
"[object Object]"`, naming a helper the caller never called and a delimiter
that was not the problem. Passing the Worksheet the other way round, as
`getCellAtAddress(ws, 'Data!A1')`, escaped as a raw `TypeError: wb.sheets is
not iterable` rather than an `OpenXmlError`. Both functions now check both
arguments and report which one is wrong, what they expected, what they
received, and, when the value is an object, point at `getCellByCoord` /
`setCellByCoord` for the case where the worksheet is already in hand.
`parseSheetRange` names the type it was handed rather than stringifying it.
