---
"@office-kit/xlsx": patch
---

Resolve `sharedStrings`, `styles` and the theme through the workbook
relationships in both loaders. `loadWorkbookStream` looked only at
`xl/sharedStrings.xml` and `xl/styles.xml`, so a workbook that keeps either
part elsewhere (legal, and what some producers emit) streamed back `null` for
every shared-string cell and resolved every style against an empty pool, with
no error to signal it. `loadWorkbook` did consult the rels, but only after the
conventional path, so an unrelated `xl/sharedStrings.xml` left in the package
shadowed the part the rels name.

A relationship that resolves to nothing is now a malformed package rather than
an absent part: both loaders throw an `OpenXmlSchemaError` when a
workbook-level relationship targets a part the package does not contain or
carries `TargetMode="External"`. Percent-encoded targets
(`Target="shared%20strings.xml"` for the entry `xl/shared strings.xml`) resolve
to the entry they name instead of reading as missing. Exact ZIP entry names
take precedence; decoding is a fallback for workbook-level optional parts,
so existing percent-encoded entry names continue to load.

`loadWorkbookStream` also rejects the malformed packages `loadWorkbook`
rejects, rather than quietly returning less than the file declares:

- a sheet whose `r:id` has no matching relationship, which used to be dropped
  from `sheetNames`
- a duplicate sheet name, where the last part won and the first became
  unreachable
- a workbook that declares sheets with no `xl/_rels/workbook.xml.rels`, which
  used to open with an empty sheet list

It reads the workbook rels whatever the sheet count now, as `loadWorkbook`
does, so a package with no sheets and an unparseable rels part fails to open
instead of opening empty.
