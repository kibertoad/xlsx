---
'@office-kit/xlsx': patch
---

fix: a part carrying the text `<!DOCTYPE` outside the prologue no longer fails to load

DTD declarations are rejected, and the check that enforced it read the whole
document. A comment or CDATA section carrying those characters as content
therefore failed the load, even though a declaration is only legal before the
root element. The check now stops once the root element opens, and matches the
whitespace the XML grammar requires after the keyword, so a read that happened
to pause right after `<!ENTITY` is no longer mistaken for one.
