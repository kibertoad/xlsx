---
"@office-kit/xlsx": patch
---

Improve failed-load diagnostics with bounded hints for text/CSV, UTF-8 and UTF-16 byte-order marks, PDF and ZIP prefixes. Empty or short inputs report their length against the 22-byte ZIP trailer minimum. ZIP errors distinguish whether an end-of-central-directory signature was found without treating that signature as proof of a complete archive.

Document the load error hierarchy, source I/O retry considerations and resource limits. Limit errors do not establish workbook validity, and callers should branch on error classes rather than message text.
