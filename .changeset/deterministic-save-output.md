---
'@office-kit/xlsx': minor
---

feat!: `SaveOptions.mtime` for byte-identical output, and `compressionLevel` is now wired through

ZIP has no "no timestamp" encoding, so fflate stamped the wall clock into every
entry's local header and central-directory record. Two saves of the same workbook
therefore differed in bytes, which ruled out golden-file tests and
content-addressed caching for anything this library writes.

`saveWorkbook` / `workbookToBytes` now accept `mtime?: Date`, applied to every
entry. `createWriteOnlyWorkbook` takes the same option. With `mtime` pinned and the
core properties set from the caller's own data rather than the clock, identical
input produces identical bytes.

The stamp records the date's UTC wall time, so the archive is the same on every
machine. ZIP's DOS date field carries no timezone, and fflate reads a `Date`
through local-time getters, which would have made the bytes depend on the writer's
`TZ`: a golden file committed from a laptop would not match the one CI renders
from the same input. A date whose year falls outside 1980-2099, the range the
field can hold, is rejected with an `OpenXmlIoError` before anything is written
rather than part-way through the first entry. Resolution is two seconds, per the
format.

`SaveOptions.compressionLevel` was declared but documented as "Reserved" and never
read; it now reaches fflate's deflate stream on both the buffered and the streaming
writer, and is typed `0 | 1 | ... | 9` rather than `number`. Two things change for
callers who already set it: the type narrows, so `{ compressionLevel: someNumber }`
no longer typechecks, and the value is now honoured, so output that silently came
back at fflate's default level 6 changes in size and in bytes. A level outside
`0..9` throws instead of falling back to fflate's default, which is what the old
"Reserved" option effectively did. `CompressionLevel` is exported from
`@office-kit/xlsx/io` and `@office-kit/xlsx/streaming` alongside the options types
that use it.
