---
"@office-kit/xlsx": patch
---

Reject a zip entry that declares DEFLATE but carries no compressed bytes,
instead of hanging on it. `loadWorkbookStream` reads worksheet parts through
`readStream`, and with no input to drive the inflater that stream never
produced a chunk, never closed and never failed, so the read span in a busy
loop that could not be cancelled. A crafted archive was enough to trigger it.
The buffered `read` path returned an empty part for the same entry and now
reports the same `OpenXmlIoError`.
