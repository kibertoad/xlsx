---
'@office-kit/xlsx': minor
---

feat!: `addExcelTable` now rejects a table that disagrees with the sheet under it

A table whose `columns` count did not match the width of its `ref`, or whose header
cells did not hold the column names it declared, produced a workbook Excel treats as
damaged. Excel "repairs" it by dropping the table, so the mistake surfaced as missing
filters and a broken structured reference in the delivered file, a long way from the
`addExcelTable` call that caused it.

Both cases now throw `OpenXmlSchemaError` at the call, naming the offending header
cell and what it holds. The header check compares rendered text rather than the raw
cell value, matching how Excel pairs a column with its header, so a rich-text or
numeric header cell spelling the right name is accepted. Write the header row before adding the table, or pass
`headerRowCount: 0` for a genuinely header-less table.

This can newly throw for code that previously appeared to work. Those are exactly the
files Excel was already repairing.
