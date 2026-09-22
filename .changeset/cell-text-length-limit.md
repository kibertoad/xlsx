---
"@office-kit/xlsx": minor
---

Writing a cell longer than Excel's 32,767-character ceiling now throws `OpenXmlSchemaError` naming the cell and both lengths, instead of producing a workbook Excel reports as needing repair and then truncates. Applies to plain strings, rich text (counted across the cell's runs, not per run) and a cached formula string result, on both the modelled and the write-only path.

Adds `MAX_CELL_TEXT_LENGTH` to `@office-kit/xlsx/utils` so a caller with text of unknown length can truncate or split before the write rather than handle a thrown error.

The count is in characters, matching Excel: an astral character is two UTF-16 code units but one character, so a cell can hold 32,767 emoji.
