---
"@office-kit/xlsx": patch
---

Accept worksheets whose `<row>` elements omit `@r`, which ECMA-376 allows.
`loadWorkbook` rejected such a file with "missing required @r", and
`loadWorkbookStream` was worse: it numbered every affected row 0, filtered it
out for sitting below the first row, and reported an empty sheet with no error
at all. Both now number an `@r`-less row as the one after its predecessor, the
same rule already applied to a cell that omits `@r`.
