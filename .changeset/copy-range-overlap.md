---
"@office-kit/xlsx": patch
---

Fix `copyRange` corrupting an overlapping same-sheet copy. The write loop read `ws.rows` as it went, so a copy shifted by less than its own height or width re-read cells it had already written: `copyRange(ws, 'A1:A3', 'A2:A4')` over `[1, 2, 3]` produced `[1, 1, 1, 1]` instead of `[1, 1, 2, 3]`. Source cells are now read in full before the first write, matching `moveRange`.
