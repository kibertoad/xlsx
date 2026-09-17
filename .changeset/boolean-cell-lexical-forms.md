---
"@office-kit/xlsx": minor
---

Read `<c t="b"><v>true</v></c>` as TRUE. Both loaders compared the value against
`"1"`, so a workbook written with the long xsd:boolean spelling (legal, and what
some non-Excel producers emit) had every TRUE cell come back as FALSE with
nothing raised to say so. All four lexical forms now read correctly in cell
values and in cached formula results, in `loadWorkbook` and
`loadWorkbookStream` alike, and the padding a pretty-printer leaves around a
value no longer changes how it reads. Boolean attributes across the worksheet,
chart and style parts go through the same parser, so they accept the same
spellings.

Three input shapes are read differently than before:

- `<v>true</v>` / `<v>false</v>` give TRUE / FALSE instead of both FALSE.
- `<c t="b"/>` reads as an empty cell rather than FALSE, and saves as an untyped
  `<c/>`: with no value there is no boolean for `t="b"` to describe, and an
  empty cell is what Excel writes.
- A value outside the lexical space (`<v>yes</v>`) makes `loadWorkbook` throw
  `OpenXmlSchemaError` naming the cell, instead of returning FALSE.
  `loadWorkbookStream` reads it as an empty cell, staying lenient the way it is
  for every other unreadable value.
