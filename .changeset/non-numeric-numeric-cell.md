---
"@office-kit/xlsx": patch
---

Report a corrupt numeric cell at load instead of at save. `<c t="n"><v>oops</v></c>`
parsed to `NaN` and an overflowing exponent to `Infinity`, and both were stored
on the cell: the load succeeded and the error surfaced much later as
`cannot serialise non-finite number` from the writer, naming a cell the caller
never wrote. `loadWorkbook` now throws an `OpenXmlSchemaError` naming the
offending text, and `loadWorkbookStream` reads the cell as empty, matching how
it handles other unreadable values.
