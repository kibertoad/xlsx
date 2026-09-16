---
'@office-kit/xlsx': minor
---

feat!: `addTable` / `addExcelTable` now reject a table that disagrees with the sheet under it

A table whose `columns` count did not match the width of its `ref`, whose `ref` left no
data row under the header, whose column names were empty or duplicated, or whose header
cells did not hold the column names it declared, produced a workbook Excel treats as
damaged. Excel "repairs" it by dropping the table, so the mistake surfaced as missing
filters and a broken structured reference in the delivered file, a long way from the call
that caused it.

All of those now throw `OpenXmlSchemaError` at the call, naming the table and the offending
header cell. A header cell has to hold text (a string, rich text, or a formula caching a
string), since that is what Excel keeps in `tableColumn/@name`: write a numeric or date
header as a string. Write the header row before adding the table, or pass
`headerRowCount: 0` for a genuinely header-less table.

`loadWorkbook` and `saveWorkbook` are unchanged: a mismatched table read from an input file
still loads, and still saves, so read-modify-write of someone else's file keeps working.

This can newly throw for code that previously appeared to work. Those are exactly the files
Excel was already repairing.
