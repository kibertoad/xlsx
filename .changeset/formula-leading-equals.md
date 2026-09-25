---
"@office-kit/xlsx": minor
---

Formula text that still begins with `=` after the leading one is removed is now rejected instead of reaching the file. Only one `=` came off, so a hand-built `{ kind: 'formula', formula: '==A1' }` was written as `<f>=A1</f>`, which Excel reports as damaged, and `makeFormula('==A1')` stored `=A1` on the value and then wrote `<f>A1</f>`, a different formula from the one the caller passed.

`makeFormula`, `setFormula`, `bindValue`, the array / shared / data-table constructors, `makeDefinedName`, `makeCfRule` and `makeDataValidation` now throw `OpenXmlSchemaError` for such text, as do the `<f>`, `<formula>`, `<formula1>`, `<formula2>` and `<definedName>` serialisers for a value built as a literal. The message names the constructor that rejected the text, or the cell being written; for `setFormula` and `bindValue` that constructor is `makeFormula`. `'==A1'` is invalid in Excel's formula bar too, so there is no reading to recover: stripping the second `=` would silently store `A1`. `bindValue` follows Excel's cell entry for the one string Excel keeps as text: a lone `'='` is now stored as text instead of throwing, and `inferCellType('=')` returns `'s'`.

Reading repairs such text instead of refusing the file: every leading `=` now comes off. A `<formula>`, `<formula1>`, `<formula2>` or `<definedName>` stored as `==$A$1>0` used to load as `=$A$1>0` and now loads as `$A$1>0`, and an `<f>` with three or more leading `=` no longer keeps one.
