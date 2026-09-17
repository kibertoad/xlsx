---
'@office-kit/xlsx': patch
---

feat: setHyperlinks and setComments, for sheets that carry one per row

`setHyperlink` and `setComment` each scan the sheet's list to find the ref they
replace, so putting a link or a note on every row costs time quadratic in the
number of rows: 16 000 hyperlinks one at a time takes about 380 ms, and 16 000
comments about 270 ms.

The new `setHyperlinks(ws, entries)` and `setComments(ws, entries)` take the
whole batch and resolve it against one index, built and dropped inside the call.
The same 16 000 hyperlinks take about 7 ms and the comments about 3 ms. Each
batch leaves the sheet exactly as the matching run of single calls would, down
to the order entries end up in.

The single-entry functions are unchanged, and no index outlives a call, so
`ws.hyperlinks` and `ws.legacyComments` stay ordinary arrays that can be read
and edited directly.
