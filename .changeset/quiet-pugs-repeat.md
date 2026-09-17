---
'@office-kit/xlsx': minor
---

perf: streamed sheet reads no longer buffer a whole chunk's parse events

Walking a worksheet queued every SAX event a single parser write produced
before yielding the first row, so peak heap tracked whatever the parser was
handed rather than staying flat. How much that was depended on the shape of
the input: a materialised 40 MB sheet body went in whole, a streamed one at
the zip reader's inflate granularity of roughly 450 KB. Both are now decoded
and fed in 64 KB slices, whatever the producer hands over, and the event queue
drains between writes.

On a 200k-row, 5-column sheet, peak heap for `iterRows({ minRow: 2 })` drops
from about 1240 MB to about 70 MB, and for a full-sheet walk from about
200 MB to about 60 MB.

Three changes in behavior come with it:

- `iterRows({ minRow })` above row 1 no longer fails on sheets written by
  Excel. The band was replayed inside a rebuilt `<sheetData>` envelope that
  declared only the default namespace, so the `x14ac:dyDescent` Excel puts on
  nearly every `<row>` raised an unbound-prefix error. The replay now carries
  the worksheet's and sheetData's own namespace declarations.
- Malformed XML throws `OpenXmlSchemaError` with the parser's own error as
  `cause`, instead of surfacing the raw `saxes` `Error`.
- Streamed input is scanned for DTD and entity declarations across the whole
  document instead of only its first 256 characters. A `<!DOCTYPE` or
  `<!ENTITY` token further into the payload, inside a comment or a CDATA
  section for instance, is now rejected where it previously parsed.
