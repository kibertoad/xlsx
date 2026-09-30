---
"@office-kit/xlsx": patch
---

fix: `loadWorkbook` dropped the sheet from a sheet-qualified defined name in the dependent cells of a shared formula. A formula `Data!total+B1` filled down read `total+B2` in the next cell, which then referred to a different name or to none, and a save wrote that text back out. It now reads `Data!total+B2`.
