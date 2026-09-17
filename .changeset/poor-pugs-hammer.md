---
'@office-kit/xlsx': minor
---

perf: saveWorkbook no longer holds whole worksheets in memory

Worksheets are serialised straight into their ZIP entry instead of being built
as a string, encoded, and held until every sheet is done. Saving a single
500k-cell sheet now completes under an 88 MB heap cap where it previously needed
192 MB; an eight-sheet, 2M-cell workbook needs 320 MB instead of 384 MB.

**Saved bytes change.** `xl/_rels/workbook.xml.rels` is now written after the
worksheet parts rather than before them, because its contents depend on whether
serialising the sheets produced any shared strings. Code that reads the package
through `loadWorkbook`, a zip library, or Excel is unaffected, since OPC resolves
parts by name and archive position carries no meaning. Byte-for-byte consumers
are: if you pinned `mtime` for reproducible output (added in 0.13.0) and store
golden files or content hashes of saved workbooks, every one of them changes
with this release and has to be regenerated.
