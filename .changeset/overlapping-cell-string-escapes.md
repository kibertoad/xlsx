---
"@office-kit/xlsx": patch
---

Fix cell text that already looks like the `_xHHHH_` escape convention losing characters on save. The writer protects such a sequence by escaping the underscore that opens it, and it missed openers in two shapes: one whose closing underscore is shared with the next sequence (`_x0041_x0042_` came back as `_x0041B`), and one whose closer is a character the writer escapes in turn, such as a line break (`"SKU_x0041\nrest"` came back as `SKUAx000A_rest`). Text of this shape now round-trips through shared strings, rich-text runs and inline strings. A workbook written by an earlier version is unaffected: it still reads back the way Excel reads it.

`loadWorkbook` now decodes each `<t>` of a shared string on its own, the way `loadWorkbookStream` already did. The two returned different text for an entry carrying more than one `<t>`, a shape the schema does not allow but both readers accept.
