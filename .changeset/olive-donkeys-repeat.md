---
'@office-kit/xlsx': minor
---

fix: reading an entry twice no longer counts it twice against the decompression budget

`readStream()` followed by `read()` on the same archive entry charged its payload
to `maxTotalUncompressedBytes` twice, so a legitimate workbook could be rejected
as a decompression bomb. Each entry is now charged the largest amount any single
inflate of it produced, so reading one again costs nothing however the reads
overlap. The per-entry size and ratio caps are unchanged, and an archive whose
distinct entries genuinely exceed the total is still rejected.

An entry rejected for exceeding the archive total remains rejected on repeated
sync or streaming reads, including archives with understated directory sizes.

perf!: loading a workbook no longer holds every part it has read

Inflated entries were cached for the lifetime of the archive, which put the whole
uncompressed package in memory. The cache now takes entries of 64 KB or less, 4 MB
of them in total, and drops the least recently used first, which still covers the
`.rels` parts a load re-reads. Loading a 10-sheet, 8.3 MB workbook holds about
21 MB of inflated bytes instead of about 65 MB.

`openZip().read(path)` returns a fresh array on every call as part of this. It
previously handed back the same array once an entry had been read, so mutating one
read's result changed what later reads of that path returned. Code that relied on
that aliasing, or on two reads yielding the identical object, has to keep its own
reference now.

Repeated row-band queries retain their worksheet bytes and row index without
re-inflating the part. Closing the workbook releases these caches and prevents
subsequent band queries from using stale bytes.
