---
"@office-kit/xlsx": minor
---

Reject a corrupt numeric cell at load instead of at save. `<c t="n"><v>oops</v></c>`
parsed to `NaN`, and an exponent past the double range to `Infinity`. Both were
stored on the cell, so the load succeeded and the failure surfaced much later as
`cannot serialise non-finite number` from the writer, naming a cell the caller
never wrote.

`loadWorkbook` and `loadWorkbookStream` both throw an `OpenXmlSchemaError` that
names the sheet, the cell and the offending text. **This changes behavior**: a
file that used to load, carrying `NaN` or `Infinity` on a cell, is now refused.
That file was already unsaveable. Ordinary numbers, an empty or whitespace-only
`<v>`, and a `<c/>` with no value are unaffected.

`saveWorkbook` also refuses a non-finite cached formula value, which it used to
write out as `<v>NaN</v>`, a part Excel cannot open.
