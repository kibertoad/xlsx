---
"@office-kit/xlsx": minor
---

**Behaviour change:** formula text and cached formula results are written
differently than in earlier versions.

Formula text (`<f>`, plus `<formula>` / `<formula1>` / `<formula2>` on
conditional formats and data validations) and a `t="str"` cached result used to
go through the cell-string escaper, which turns a literal `_x0041_` into
`_x005F_x0041_`. Nothing decodes those nodes on read, so `CONCAT("_x0041_")`
was already wrong in the first saved file and grew another `_x005F_` on every
load-and-save cycle after that. They now take plain XML escaping and survive
any number of cycles unchanged.

Along the same paths:

- A carriage return is written as `&#13;`, so it comes back as a CR instead of
  being normalised to a line feed.
- A codepoint XML 1.0 cannot represent (a C0 control character other than tab /
  LF / CR, an unpaired surrogate, or U+FFFE / U+FFFF) now throws `OpenXmlSchemaError` naming the
  cell, rather than being encoded as an `_xHHHH_` sequence nothing reverses.
- A cached error result keeps `t="e"` instead of being downgraded to `t="str"`,
  so `ISERROR` / `IFERROR` and error-keyed conditional formats still match it
  before Excel recalculates. To create one, pass `cachedValueType: 'error'`
  alongside a cached error token to the formula constructor. Plain cached
  strings, including `"#N/A"`, stay strings; loaded errors retain their type.
- A cached result stored in the shared-strings table (`t="s"`, which non-Excel
  producers write) resolves to its text. It used to load as the raw sst index
  and save back as that number.

Formulas in files written by earlier versions keep whatever `_x005F_` prefixes
they accumulated, since a stored `_x005F_x0041_` is indistinguishable from one
Excel wrote on purpose. They no longer grow.
