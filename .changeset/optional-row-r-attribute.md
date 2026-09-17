---
"@office-kit/xlsx": patch
---

Accept worksheets whose `<row>` elements omit the optional `r` attribute, which
ECMA-376 allows. `loadWorkbook` rejected such a file with "missing required @r",
and `loadWorkbookStream` numbered every affected row 0, filtered it out for
sitting below the first row, and reported an empty sheet with no error at all.

Both readers now place such a row where its first located cell says, or, with no
cell to go by, on the row after the highest one read so far. A streaming band
query (`minRow` / `maxRow`) covers those rows: a sheet that omits `r` cannot be
seeked into by row number, so band queries stream it instead of jumping to a
byte offset.

`<row r="…">` values that are not a row number in `[1, 1048576]` now throw an
`OpenXmlSchemaError` from both readers; `loadWorkbookStream` used to drop such a
row silently, and `loadWorkbook` used to accept a value past the last row.

A located cell can appear after unlocated cells in the same row. Both readers
assign all of them to the derived row before applying a row band. Unlocated
column numbering also stays independent of column filters. Row attributes
continue to accept the optional plus sign allowed by `xsd:unsignedInt`.
