---
"@office-kit/xlsx": patch
---

Fix a hang when loading a corrupt archive. A zip entry that declared DEFLATE
but carried no compressed bytes sent `loadWorkbookStream` into a busy loop that
could not be cancelled, and a crafted file was enough to trigger it. Both read
paths now reject that entry with an `OpenXmlIoError`; the buffered `read` path
previously returned an empty part.

`loadWorkbook` and `loadWorkbookStream` also reject an entry whose declared
compressed size runs past the end of the archive, instead of reading whatever
bytes were left. That case previously loaded, with binary parts such as
`vbaProject.bin` or images silently truncated and written back out on save.

Both rejections fail the whole load, so an archive carrying one corrupt part
that earlier releases read past now throws.
