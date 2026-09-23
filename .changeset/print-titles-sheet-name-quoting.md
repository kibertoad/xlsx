---
"@office-kit/xlsx": patch
---

Fix `setPrintTitles` producing an unreadable reference when the sheet name contains an apostrophe. It wrapped the name in quotes without doubling the apostrophes inside it, so a sheet called `Bob's Sheet` yielded `'Bob's Sheet'!$1:$1`, which Excel treats as a broken defined name and which `getDefinedNameTarget` rejected with a missing-delimiter error. Titles with apostrophes are accepted by `validateSheetTitle`, so this was reachable from supported input.
