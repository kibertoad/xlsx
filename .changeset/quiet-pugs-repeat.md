---
'@office-kit/xlsx': patch
---

perf: streamed sheet reads no longer buffer the whole sheet's parse events

Walking a worksheet through `loadWorkbookStream` queued every SAX event the
sheet produced before yielding the first row, so peak heap scaled with the
sheet instead of staying flat. A 40 MB sheet body peaked at about 975 MB of
heap; it now peaks at about 90 MB.

`iterRows({ minRow })` with `minRow` above 1 was hit hardest, because that path
hands the parser a materialised buffer. On a 200k-row sheet, `{ minRow: 2 }`
went from about 1640 MB peak heap to about 105 MB.
