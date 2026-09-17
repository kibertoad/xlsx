---
"@office-kit/xlsx": patch
---

Stop rewriting formula text that contains a literal `_xNNNN_`. The writer ran
`<f>` bodies and `t="str"` cached results through the cell-string escaper,
which turns `_x0041_` into `_x005F_x0041_`. Excel does not use that convention
in formulas and nothing unescapes them on read, so `CONCAT("_x0041_")` was
already wrong in the first saved file and grew another `_x005F_` on every
load-and-save cycle after that. Both now take plain XML escaping.

Formulas in files written by earlier versions keep whatever `_x005F_` prefixes
they accumulated, since a stored `_x005F_x0041_` is indistinguishable from one
Excel wrote on purpose. They no longer grow.
