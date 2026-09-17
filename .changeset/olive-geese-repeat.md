---
'@office-kit/xlsx': patch
---

fix: a failed `saveWorkbook` to a file path no longer leaves the partial file behind

`toFile` deletes the half-written file when a save fails, but the delete raced
the write stream's own file-handle close and lost on Windows, so a truncated
`.xlsx` stayed at the destination. The cleanup now waits for the handle to
close, and `saveWorkbook` (and the write-only `finalize()`) wait for the cleanup
before rejecting, so the path is clear by the time the error reaches the caller.
