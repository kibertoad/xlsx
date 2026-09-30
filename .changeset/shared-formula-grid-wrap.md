---
"@office-kit/xlsx": patch
---

fix: `loadWorkbook` failed on a shared formula filled into column XFD, and gave a shared formula filled down to the last row a reference past row 1048576. Both now wrap to the opposite edge of the grid, which is what Excel shows for the same file: a formula referencing `B1048576`, filled one row down, reads `B1`. A token past the grid such as `A2000000` is a name to Excel and is no longer shifted.
