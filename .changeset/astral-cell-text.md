---
"@office-kit/xlsx": patch
---

Fix emoji and other astral-plane characters being written as `_xD83D__xDE00_`.
Cell strings, rich-text runs, formula text and cached formula results all
escaped the two halves of a surrogate pair separately, so Excel, LibreOffice
and Google Sheets displayed the escape sequence instead of the character.
Reading such a file back through this library hid the problem, because the
unescape step rebuilt the character. Unpaired surrogates, which have no UTF-8
encoding, are still escaped, and files already written with the split form
still read back correctly.
