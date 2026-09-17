---
'@office-kit/xlsx': patch
---

fix: XML containing literal `<!DOCTYPE` or `<!ENTITY` in comments, processing instructions, or CDATA no longer fails to load

The declaration check now respects XML lexical context and stops at the root
start tag, even within a chunk. Valid content loads consistently from strings,
bytes, and streams regardless of chunk boundaries. Actual DTD and entity
declarations remain forbidden.
