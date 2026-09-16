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

`SaveOptions.compressionLevel` was declared but documented as "Reserved" and never
read; it now reaches fflate's deflate stream on both the buffered and the streaming
writer, and is typed `0 | 1 | ... | 9` rather than `number`. That narrowing is the
only breaking part of this change: `{ compressionLevel: someNumber }` no longer
typechecks, though it never did anything either.
