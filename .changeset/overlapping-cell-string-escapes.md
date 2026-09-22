---
"@office-kit/xlsx": patch
---

Fix cell text that already looks like the `_xHHHH_` escape convention losing characters on save. Two such sequences can share an underscore (`_x0041_x0042_`), and the writer's protection consumed the one it matched, leaving the following sequence exposed for the reader to decode: the cell came back as `_x0041B`. Strings of the affected shape now round-trip, and a workbook already written with the old escaping still reads back the way Excel reads it.
