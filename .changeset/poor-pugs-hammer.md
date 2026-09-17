---
'@office-kit/xlsx': patch
---

perf: saveWorkbook no longer holds whole worksheets in memory

Worksheets are serialised straight into their ZIP entry instead of being built
as a string, encoded, and held until every sheet is done. Saving a single
500k-cell sheet now completes under an 88 MB heap cap where it previously needed
192 MB; an eight-sheet, 2M-cell workbook needs 320 MB instead of 384 MB.

One byte-level detail changed: `xl/_rels/workbook.xml.rels` is now written after
the worksheet parts rather than before them. Its contents depend on whether
serialising the sheets produced any shared strings, and OPC resolves parts by
name, so archive position carries no meaning. Code that reads the package
through `loadWorkbook`, a zip library, or Excel is unaffected; only a byte-for-byte
comparison against output from an earlier version would notice.
