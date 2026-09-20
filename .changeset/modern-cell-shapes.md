---
"@office-kit/xlsx": minor
---

Load the cell shapes a current Excel writes. A `<c t="e">` carrying one of the
nine error tokens Excel added since 2018 (`#SPILL!`, `#CALC!`, `#FIELD!`,
`#BLOCKED!`, `#CONNECT!`, `#BUSY!`, `#UNKNOWN!`, `#PYTHON!`, `#EXTERNAL!`) threw
`unknown error code`, and a `<c t="d">` holding an ISO 8601 date threw
`unknown cell type t="d"`. One such cell anywhere in a workbook aborted the
whole load, so a file with thousands of good rows failed over a single spilled
formula.

`ERROR_CODES` now lists all seventeen tokens, and a `t="e"` token outside that
set is kept verbatim rather than dropped: `t="e"` is the file declaring the cell
an error, Excel keeps adding tokens, and a load then save re-emits the token
unchanged, on a formula cell's cached result as well as a plain one. A token is
anything shaped like one (`#`, a letter, then upper-case ASCII, digits, `_` or
`/`, optionally closed by `!` or `?`); a `t="e"` holding anything else is
reported against the cell it sits in, on read and on save alike. A `t="e"` whose
`<v>` is missing or blank is an empty cell, the way it is under `t="n"` and
`t="b"`.

ISO `t="d"` cells in Transitional producer output now read as a `Date`, or
as `{ kind: 'duration', ms }` for time-only and day/time duration forms.
Unqualified calendar dates use UTC; explicit date offsets are applied and
validated within ±14:00. Time-only offsets are validated but retain the stated
wall-clock time. Fractional seconds in dates and times are truncated to
milliseconds. `24:00:00` represents one day. Invalid dates, incomplete durations
and durations beyond the safe integer millisecond range fail during loading.
Saving emits numeric serials under the workbook epoch. ISO formula caches also
become numeric serials, preserving their value with either the 1900 or 1904 epoch.

Strict packages keep the normalization contract introduced in 0.20.0: supported
calendar dates become numeric serials, while time-only/duration cells and
sub-millisecond dates remain unsupported. This does not widen that Strict scope.

**This changes `ExcelErrorCode`** from a union of eight literals to
`` `#${string}` ``. Code that passes the type around, or into `makeErrorValue`,
is unaffected; an exhaustive `switch` over the eight literals in your own code
no longer type-checks as exhaustive.

**The nine new tokens also change what `bindValue` writes.** `bindValue` stores
any string listed in `ERROR_CODES` as an error value, so
`bindValue(cell, '#PYTHON!')` now saves `t="e"` where it used to save the
literal text as a string. `makeErrorValue` accepts the nine as well. The
explicit setters (`setCellValue`, `setCell`) are unaffected.

`loadWorkbookStream` answers all of these the way `loadWorkbook` does for a
cell holding a value. It previously read an unlisted error token as an empty
cell and ran a `t` outside `ST_CellType` through `Number.parseFloat`, so the
same file could load as empty cells in one reader and throw in the other. A
formula cell still reads as its cached value there, since the streaming reader
models no formulas.
