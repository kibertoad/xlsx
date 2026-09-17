---
"@office-kit/xlsx": minor
---

Fix emoji and other astral-plane characters being written as `_xD83D__xDE00_`.
Cell strings, rich-text runs, formula text and cached formula results all
escaped the two halves of a surrogate pair separately, so Excel, LibreOffice
and Google Sheets displayed the escape sequence instead of the character.
Reading such a file back through this library hid the problem, because the
unescape step rebuilt the character.

U+FFFE and U+FFFF are now escaped too. XML 1.0 leaves both out of its `Char`
production, so a cell carrying either one produced a part that strict parsers
and Excel reject.

Unpaired surrogates, which have no UTF-8 encoding, are still escaped. A file
already written with the split form reads back correctly wherever the reader
inverts the `_xHHHH_` convention, which covers cell text and rich-text runs.
Formula text and a cached string result are handed back as written, so an emoji
split there stays split.

This changes the output of `escapeCellString` (exported from
`@office-kit/xlsx/utils`) for the inputs above. Code that diffs writer output
against a stored file containing `_xD83D_` will see it change.
