---
"@office-kit/xlsx": patch
---

fix: a failed load now says what the bytes were, and documents which error classes mean "not an xlsx"

Uploads get validated by file extension, so a CSV saved as `.xlsx` reaches
`loadWorkbook` and used to be told only `openZip: archive is not a valid zip`.
Where a magic number identifies the input, the message now names it: plain text
or a CSV, a UTF-8 or UTF-16 byte-order mark, a PDF, or a file that begins with
zip entries. That last one reads two ways, and the repair differs: entries with
no trailer behind them is what a truncated or partially uploaded file looks
like, while a trailer that is present over an unreadable central directory is a
whole file that got corrupted. An empty or short file reports its own length
against the 22-byte minimum.

`loadWorkbook`, the `OpenXmlError` classes and the README now also state the
error contract: which class means "reject the upload and tell the user" as
against "retry", that `OpenXmlContentLimitError` extends `OpenXmlError`
directly and so escapes a catch ladder built from the other classes, and that
the class is stable while the message text is not.
