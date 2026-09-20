---
"@office-kit/xlsx": minor
---

Stop telling the owner of a legacy `.xls` to decrypt it. Any OLE compound file was
reported as `Encrypted xlsx is not supported. Decrypt with msoffcrypto-tool first.`,
and a BIFF `.xls` is an OLE compound file too, so passing one to `loadWorkbook`
sent people looking for a password that was never set.

`loadWorkbook`, `loadWorkbookStream` and `openZip` now look at the streams in the
container's root storage: an `EncryptedPackage` stream keeps the decrypt advice,
and a `Workbook` or `Book` stream gets an error saying the file is a legacy `.xls`
and has to be converted to `.xlsx` first. A container that fits neither (another
Office format, or a truncated file) gets a message naming both possibilities. A
`.xls` from Excel 4.0 or earlier, which has no compound file around it, gets the
convert advice too; it used to fail as an invalid zip.

All of these throw the new `OpenXmlUnsupportedFormatError`, exported from
`@office-kit/xlsx/utils`. Its `format` field is `'encrypted-xlsx'`, `'legacy-xls'`
or `'compound-file'`, so an upload handler can ask for a password in one case and
for a re-save in the other:

```ts
try {
  await loadWorkbook(source);
} catch (error) {
  if (error instanceof OpenXmlUnsupportedFormatError && error.format === 'legacy-xls') {
    // ask for an .xlsx, or hand the file to a BIFF reader
  }
  throw error;
}
```

The class extends `OpenXmlNotImplementedError`, so existing `catch` paths still
see it. Branch on `format`, because the message wording can change. Code that
already matches the message against `Encrypted xlsx` still matches for an
encrypted file and no longer matches for the other two.
