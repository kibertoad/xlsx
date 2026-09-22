---
"@office-kit/xlsx": minor
---

`loadWorkbookStream` now reports rich text the way `loadWorkbook` does. It flattened a shared or inline string built from `<r>` runs into the joined plain text, so per-run formatting (bold, colour, size) was reachable through one entry point and not the other, even though `ReadOnlyCell.value` is typed `CellValue` and so already admitted the rich-text variant. A cell that `loadWorkbook` reports as `{ kind: 'rich-text', runs }` now reads the same way when streamed.

Callers that relied on a streamed rich-text cell arriving as a `string` should pass it through `cellValueAsString` from `@office-kit/xlsx/cell`.

The remaining difference between the two readers, which is deliberate, is documented on `loadWorkbookStream`: the streaming reader reports a cell as empty for a `t="s"` index with no shared string behind it and for a `t="b"` value outside `xsd:boolean`, where `loadWorkbook` rejects the file.
