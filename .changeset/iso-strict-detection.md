---
"@office-kit/xlsx": minor
---

feat: a workbook saved as "Strict Open XML Spreadsheet" is now named as unsupported instead of failing with a missing-relationship error

Excel's Save As dialog offers that entry and it writes a file with the `.xlsx`
extension, so a strict package reaches the readers looking like any other
workbook. Every part inside it uses the ISO 29500 strict namespace family
(`purl.oclc.org`) rather than the transitional one
(`schemas.openxmlformats.org`), which the reader is built on, and the load
failed with `OpenXmlSchemaError: loadWorkbook: root rels missing officeDocument
relationship`. `loadWorkbook` and `loadWorkbookStream` now throw
`OpenXmlNotImplementedError` naming the format and saying to re-save the file
as "Excel Workbook (.xlsx)". Converter output that mixes the two families is
caught per part, including the shape that used to load as a workbook with no
sheets and no error.

**Behavior change, flagged pre-1.0:** a strict package raises
`OpenXmlNotImplementedError` where it previously raised `OpenXmlSchemaError`,
and those two are siblings rather than one extending the other. Code that
branches on `instanceof OpenXmlSchemaError` to turn a bad upload into a
4xx stops catching strict files; catch `OpenXmlError`, or add an
`OpenXmlNotImplementedError` arm.

`loadWorkbook` also now requires `xl/workbook.xml` to declare the
SpreadsheetML namespace on its root element, not just the local name
`workbook`. A package that declared some other namespace previously loaded as
a workbook with no sheets; it now throws `OpenXmlSchemaError`.

Reading strict packages is still unimplemented. Writing is unchanged and stays
transitional.
