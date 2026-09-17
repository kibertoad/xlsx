---
'@office-kit/xlsx': patch
---

fix: reading an entry twice no longer counts it twice against the decompression budget

`readStream()` followed by `read()` on the same archive entry charged its payload
to `maxTotalUncompressedBytes` twice, so a legitimate workbook could be rejected
as a decompression bomb. The budget now tracks what each entry has contributed
and refunds it before counting a re-read. The per-entry size and ratio caps are
unchanged, and an archive whose distinct entries genuinely exceed the total is
still rejected.

perf: loading a workbook no longer holds every part it has read

Inflated entries were cached for the lifetime of the archive, which put the whole
uncompressed package in memory. The cache is now limited to entries of 64 KB or
less and 4 MB in total, which still covers the `.rels` parts a load actually
re-reads. Loading a 10-sheet, 8.3 MB workbook holds about 21 MB of inflated bytes
instead of about 65 MB.
