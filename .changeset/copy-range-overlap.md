---
"@office-kit/xlsx": minor
---

fix: `copyRange` corrupted an overlapping same-sheet copy, and both range movers left stale cells behind.

`copyRange`'s write loop read `ws.rows` as it went, so a copy shifted by less than its own height or width re-read cells it had already written: `copyRange(ws, 'A1:A3', 'A2:A4')` over `[1, 2, 3]` produced `[1, 1, 1, 1]` instead of `[1, 1, 2, 3]`. Source cells are now read in full before the first write.

Behaviour changes that go with it:

- `copyRange` / `moveRange` replace the landing rectangle whole, the way pasting over a selection in Excel does. A coordinate whose source cell is empty now lands empty instead of keeping whatever the destination held. Previously `copyRange(ws, 'A1:A3', 'C1:C3')` with an empty `A2` left `C2` untouched, so a copied block could arrive with a foreign value wedged into its gap.
- A destination cell no longer keeps its own `hyperlinkId` / `commentId` when a copy or move lands on it. Copying an unlinked cell over a linked one used to leave the new value wearing the old link. Cross-sheet copies still drop the source's ids, since those index the source sheet's own `hyperlinks` / `legacyComments` arrays.
- The return value counts populated source cells only. For a sparse source overlapping its target it can now be lower than before: the old count included destination cells the loop had created itself and then re-read.

`opts.targetWs` has to be a sheet in the same workbook. `styleId` indexes that workbook's `cellXfs` and is copied verbatim, so cells copied into a second workbook arrive wearing whichever style occupies the same slot there. This was always true; it is now stated on both functions.
