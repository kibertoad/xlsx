---
'@office-kit/xlsx': patch
---

perf: streamed sheet reads no longer buffer a whole chunk's parse events

Walking a worksheet queued every SAX event a single parser write produced
before yielding the first row, so peak heap tracked whatever the parser was
handed rather than staying flat.

How much was handed over depended on the shape of the input. A materialised
40 MB sheet body went in whole and peaked around 975 MB of heap. A streamed one
went in at the zip reader's inflate granularity, around 450 KB per chunk, and
peaked around 220 MB. Both are now fed in 64 KB chunks and peak around 65 MB
and 95 MB respectively.

For `loadWorkbookStream`, `iterRows({ minRow })` above row 1 gains the most,
going from about 1640 MB peak heap to about 105 MB on a 200k-row sheet. A
default full-sheet walk gains less, about 450 MB to about 380 MB, because what
is left there is the inflate buffers and the row objects themselves rather than
queued events.
