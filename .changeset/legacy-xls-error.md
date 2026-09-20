---
"@office-kit/xlsx": patch
---

Stop telling the owner of a legacy `.xls` to decrypt it. Any OLE compound file was
reported as `Encrypted xlsx is not supported. Decrypt with msoffcrypto-tool first.`,
and a BIFF `.xls` is an OLE compound file too, so passing one to `loadWorkbook`
sent people looking for a password that was never set.

`loadWorkbook`, `loadWorkbookStream` and `openZip` now read the container's
directory and pick the message from what it holds: an `EncryptedPackage` stream
keeps the decrypt advice, and a `Workbook` or `Book` stream gets an error saying
the file is a legacy `.xls` and has to be converted to `.xlsx` first. A container
that fits neither (another Office format, or a truncated file) gets a message
naming both possibilities. The error class is `OpenXmlNotImplementedError` in all
three cases, as before. Code that matches the message against `Encrypted xlsx`
still matches for an encrypted file and no longer matches for the other two.
