---
"@office-kit/xlsx": patch
---

fix: a workbook saved as "Strict Open XML Spreadsheet" is now named as unsupported instead of failing with a missing-relationship error

Excel's Save As dialog offers that entry and it writes a file with the `.xlsx`
extension, so a strict package reaches `loadWorkbook` looking like any other
workbook. Every part inside it uses the ISO 29500 strict namespace family
(`purl.oclc.org`) rather than the transitional one
(`schemas.openxmlformats.org`), which the reader is built on, and the load
failed with `OpenXmlSchemaError: loadWorkbook: root rels missing officeDocument
relationship`. It now throws `OpenXmlNotImplementedError` naming the format and
saying to re-save the file as "Excel Workbook (.xlsx)".

Reading strict packages is still unimplemented. Writing is unchanged and stays
transitional.
