---
"@office-kit/xlsx": patch
---

`makeFormula`, `setFormula`, `bindValue` and the other formula constructors now reject text that still begins with `=` after the leading one is removed, instead of writing a formula Excel reports as damaged. `normalizeFormulaText` strips exactly one `=`, so `'==A1'` used to be stored as `=A1` and emitted as `<f>=A1</f>`, which is the shape the function exists to keep out of the file. Such input throws `OpenXmlSchemaError` naming the text.
