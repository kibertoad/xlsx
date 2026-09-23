---
"@office-kit/xlsx": minor
---

Writing a cell past Excel's 32,767-unit ceiling now throws `OpenXmlSchemaError` naming the cell and both lengths, instead of producing a workbook Excel reports as needing repair. Applies to plain strings, rich text (counted across the cell's runs, not per run) and a cached formula string result, on both the modelled and the write-only path.

The ceiling counts UTF-16 code units, the unit `String.length` counts and Excel's own `LEN` reports: 16,384 emoji are 16,384 characters but 32,768 code units, and Excel refuses that cell.

Adds `MAX_CELL_TEXT_LENGTH` to `@office-kit/xlsx/utils` so a caller with text of unknown length can shorten it before the write rather than handle a thrown error. A slice has to keep a surrogate pair together, since a lone surrogate is rejected further down as unrepresentable in XML.

Reading is unchanged, so a file from a producer that does not check still loads; saving it back out fails until the cell named in the error is shortened.
