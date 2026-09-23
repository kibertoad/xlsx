---
"@office-kit/xlsx": minor
---

Formula text that still begins with `=` after the leading one is removed is now rejected instead of reaching the file. `normalizeFormulaText` strips exactly one `=`, so a hand-built `{ kind: 'formula', formula: '==A1' }` was written as `<f>=A1</f>`, which Excel reports as damaged, and `makeFormula('==A1')` stored `=A1` on the value and then wrote `<f>A1</f>`, a different formula from the one the caller passed.

`makeFormula`, `setFormula`, `bindValue`, the array / shared / data-table constructors, `makeDefinedName`, `makeCfRule` and `makeDataValidation` now throw `OpenXmlSchemaError` for such text, as do the `<f>`, `<formula>`, `<formula1>`, `<formula2>` and `<definedName>` serialisers for a value built as a literal. The message names the call or the cell it came from. `'==A1'` is invalid in Excel's formula bar too, so there is no reading to recover: stripping the second `=` would silently store `A1`.

Reading is unaffected. A file whose formula text carries the doubled prefix still loads, with the prefix removed, as before.
