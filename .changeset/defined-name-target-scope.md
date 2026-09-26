---
"@office-kit/xlsx": patch
---

`getDefinedNameTarget` no longer throws `missing "!" delimiter` on a sheet-scoped name whose value has no sheet prefix, such as `_xlnm.Print_Area` = `$A$1:$E$20`. Other tools, and earlier versions of `setPrintArea`, write print areas that way. An unqualified leg now resolves to the sheet the name is scoped to. A workbook-scoped name with an unqualified value still throws, since it names no sheet.
