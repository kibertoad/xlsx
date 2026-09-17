---
"@office-kit/xlsx": patch
---

`loadWorkbook` reads a large sheet about twice as fast and with under half the
transient heap. On a 50 000-row, six-column sheet (1 282 KiB archive) the load
goes from about 1 860 ms to about 850 ms, and peak RSS above the pre-load
baseline from about 840 MB to about 340 MB.

`<sheetData>` used to be read from a node tree, so every `<c>` and every `<v>`
became an object that existed only long enough to produce one cell. It is now
walked with a synchronous SAX pass; the rest of the worksheet keeps the node
tree it had.

Malformed XML inside `<sheetData>` is refused as a result, which brings
`loadWorkbook` into line with `loadWorkbookStream`. An unclosed `<row>`, a stray
`</c>`, a `<sheetData>` that is never closed, an undefined entity reference such
as `&nbsp;`, and the literal `]]>` in cell text used to be accepted, loading
whatever cells the tolerant parse happened to recover. They now throw an
`OpenXmlSchemaError`, as they already did under `loadWorkbookStream`. The rest
of the part is still read leniently, so the same undefined entity reference
outside `<sheetData>` keeps loading as the literal text it did before.

An element inside a cell's `<v>` or `<f>` is refused too, naming the cell it sat
in. It used to be dropped, leaving the element's text as the cell's value, which
is the wrong-value outcome `loadWorkbook` prefers a failed load to.

Well-formed parts, including those written with a namespace prefix and those
carrying comments or CDATA sections inside `<sheetData>`, are unaffected.
