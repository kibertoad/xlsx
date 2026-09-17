---
"@office-kit/xlsx": patch
---

Read `<c t="b"><v>true</v></c>` as TRUE. Both loaders compared the value
against `"1"`, so a workbook written with the long xsd:boolean spelling (legal,
and what some non-Excel producers emit) had every TRUE cell silently come back
as FALSE. All four lexical forms now work in cell values and in cached formula
results. A `t="b"` cell with no value reads as empty, and `loadWorkbook`
rejects a value outside the lexical space instead of quietly calling it FALSE.
