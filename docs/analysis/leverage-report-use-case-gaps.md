# Use-case analysis: TM leverage report renderer (lokalise/autopilot#9494)

Analysis of how [lokalise/autopilot#9494](https://github.com/lokalise/autopilot/pull/9494) uses
`@office-kit/xlsx@0.11.1`, what caused friction, and what the library could add or change to serve
this kind of consumer better. No code changes are proposed here; this is input for scoping.

## The use case

A pure `payload -> Uint8Array` renderer that builds a five-sheet customer-facing report from
scratch. The workbook is never read back in production; it is handed to a customer, who hands it
to a vendor, who opens it in Excel, Google Sheets or LibreOffice.

Per sheet the renderer does the same handful of things:

- write a title and a few muted notes (italic, small, grey)
- write a bold, filled, bordered, wrap-aligned header row and freeze above the first data row
- write data rows with thousands-separated integers and a thin border on every cell
- put live formulas (`SUMIFS`, `SUM`, `IF(OR(...))`, `n*$B$5`) in specific cells, with a number
  format, no cached value
- an autofilter over the tidy data range on one sheet
- fixed column widths
- `creator`, `created` and `modified` stamped from the payload, not the clock

Tests load the produced bytes back with `loadWorkbook` and assert on sheet names, cell values,
formula text, autofilter ref and core properties.

### Library surface actually used

| Subpath      | Functions                                                                                                                                                  |
| ------------ | ---------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `/cell`      | `setFormula`, `getFormulaText`, types `Cell`, `CellValue`                                                                                                   |
| `/io`        | `workbookToBytes`, `loadWorkbook`, `fromArrayBuffer`                                                                                                        |
| `/styles`    | `makeFont`, `makeColor`, `makePatternFill`, `makeSide`, `makeBorder`, `makeAlignment`, `getCellFont`, `setCellFont`, `setCellFill`, `setCellBorder`, `setCellAlignment`, `setCellNumberFormat` |
| `/utils`     | `columnLetterFromIndex`                                                                                                                                     |
| `/workbook`  | `createWorkbook`, `addWorksheet`, `getSheet`, `sheetNames`, `Workbook.properties`                                                                           |
| `/worksheet` | `setCell`, `getCell`, `iterCells`, `setColumnWidths`, `setFreezePanes`, `makeAutoFilter`, `setAutoFilter`, `getAutoFilter`                                  |

Everything the renderer needed was reachable. The friction is in how many calls it took, and in
two behaviours that surprised the author badly enough to earn a paragraph in the PR description.

## Friction observed in the PR

The PR text and the code are candid about where the library got in the way. Quoted or paraphrased
from the PR description and the renderer's own comments:

1. **`setCell` without a value blanks the cell.** "The first port of this renderer lost 34 of its
   37 formulas that way, because the border pass ran over already-populated rows." The renderer
   now carries a `put` / `cellAt` split, where `cellAt` is `getCell(...) ?? setCell(ws, row, col)`.
2. **Replacing a font drops the workbook default.** The renderer wraps every font change in a
   `patchFont` helper that spreads `getCellFont` before `makeFont`, "so assigning
   `makeFont({ bold: true })` outright would also drop the workbook default (Calibri 11)".
3. **Per-cell styling loops.** Each data row runs one loop to stamp `#,##0` on three columns and a
   second loop to stamp a border on eight columns, so a 16-row sheet makes roughly 200 style calls.
   The header helper makes four style calls per header cell.
4. **Formula placement is two steps.** Every formula is `setFormula(cellAt(sheet, r, c), text)`
   followed by a separate number-format call.
5. **A1 strings from numeric coordinates.** The renderer imports `columnLetterFromIndex` to build
   `SUM(B5:I5)`, `A4:H20` and `$B$5`, because it tracks rows and columns as numbers and the
   formula and range APIs want letters.
6. **Non-deterministic bytes.** The determinism test compares parsed cells rather than bytes,
   because "the xlsx container stores a DOS mtime per zip entry, taken from the wall clock".
7. **Loading raw bytes in tests.** `loadWorkbook` needs an `XlsxSource`, and the helper that
   wraps a `Uint8Array` is called `fromArrayBuffer`. The test comments on this.
8. **Hand-rolled inspection helpers.** `sheetOrThrow`, `collectFormulas` and `describeCells` are
   written locally to assert on the produced workbook.

Some of these already have an answer in the library that the author did not find. Those are
documentation gaps rather than API gaps, and are called out as such below.

## Gaps and improvements

Ordered by how much they would have changed this PR.

### 1. `setCell` blanking is a trap; add an explicit get-or-create and make the value mandatory

`setCell(ws, row, col, value = null)` writes `null` when the fourth argument is omitted. The
docstring says "Create or update a Cell" and promises that style, hyperlink and comment ids
survive, but never says the value is replaced. The cheatsheet shows `setCell(ws, 3, 1, null)` and
the formula recipe shows `setFormula(setCell(ws, 4, 1), ...)`, so the docs actively teach the
no-value form as the way to reach a cell.

Internally the library already has the pattern the PR had to write: `applyToRange`,
`setRangeStyle` and `setRangeBorderBox` all do `ws.rows.get(r)?.get(c) ?? setCell(ws, r, c)`.
That get-or-create is not public.

Options:

- **Export a get-or-create helper** (`ensureCell(ws, row, col): Cell`) and have the internal
  callers use it. The PR's `cellAt` becomes a one-line import.
- **Make `value` a required parameter of `setCell`.** With `ensureCell` available, the no-value
  form of `setCell` has no legitimate use, and requiring the argument turns the footgun into a
  compile error. This is a type-level break for anyone calling `setCell(ws, r, c)` today, so it
  needs a changeset and a note, but pre-1.0 this is the moment to do it.
- **Document the semantics** on `setCell`, in the cheatsheet and in the formula recipe (which
  should use `ensureCell` or pass the formula value directly).

Recommendation: all three. Keeping `setCell` optional-value and only adding docs leaves the trap
armed for the next report renderer.

### 2. One-call multi-field font patch

`setBold`, `setItalic`, `setFontSize`, `setFontName` and `setFontColor` already patch a single
field while preserving the rest, which is what `patchFont` does. The author did not find them, and
even with them, "bold + size 13" or "bold + red" costs two calls and two intermediate xf records.

The deeper problem is that `setCellFont(wb, c, makeFont({ bold: true }))` is a legal call that
registers a `<font><b/></font>` with no `<sz>` or `<name>`. Excel falls back to its own default,
LibreOffice and Sheets may render differently, and nothing warns.

Options:

- **Add `patchCellFont(wb, c, patch: Partial<Font>)`** that merges over the current font. Rebuild
  the five single-field setters on top of it so there is still one merge path.
- Same shape for the other axes is not needed: fill, border and number format are replace-whole
  by nature, and `setCellAlignment` already takes a complete object that `makeAlignment` fills.
- Consider having `setRangeStyle` / `setCellStyle` accept `font: Partial<Font>` and merge, since
  those are the calls a report renderer reaches for once it stops styling cell by cell.

Documentation: the cheatsheet row "Bold + font size + fill on a header cell" lists the right
functions; a "muted note" row and a "patch several font fields" row would have caught this case.

### 3. Build a style once, apply it by id

`setCell` already accepts a `styleId`, and the stylesheet dedups xf records, but there is no public
way to obtain a `styleId` from a style spec without first having a cell. `setRangeStyle` builds
exactly that xf internally and then throws the id away.

A report renderer wants:

```ts
const HEADER = registerCellStyle(wb, { font, fill, border, alignment });
const INT_BORDERED = registerCellStyle(wb, { numberFormat: '#,##0', border: THIN_BORDER });
setCell(ws, r, c, value, INT_BORDERED);
```

This replaces the two per-row loops and the four-call header helper with one argument per write.
It also composes with `appendRow` if `appendRow` grows an optional per-column `styleIds` array.

Proposal: export a `registerCellStyle(wb, opts): number` (the same `opts` shape as `setCellStyle`)
and have `setCellStyle` / `setRangeStyle` use it. The one-way rule holds: `setCellStyle` remains
the way to style an existing cell, `registerCellStyle` is the way to get an id for a write.

Caveat: a style built from scratch does not inherit a cell's existing xf, so the docs must say the
id is a full style, not a patch.

### 4. Range helpers should accept numeric bounds

`setRangeStyle`, `setRangeNumberFormat`, `setRangeBorderBox`, `applyToRange`, `writeRange` and
`makeAutoFilter` all take an A1 string. A renderer that walks rows and columns as integers has to
go through `columnLetterFromIndex` or `boundariesToRangeString` (which exists in `/utils` but is
not mentioned anywhere a report author would look).

Options, in order of preference:

- Accept `string | CellRangeBoundaries` on every `range` parameter. The parse step already
  produces `CellRangeBoundaries`, so the change is a type union plus an early return.
- Or promote `boundariesToRangeString` and a `rangeRef(minRow, minCol, maxRow, maxCol)` shorthand
  in the cheatsheet next to every range helper.

The same applies to formula text. A small `cellRef(row, col, { absoluteRow, absoluteCol })` helper
would have removed the `$B$${rowNumber}` and `${totalLetter}${firstDataRow + index}` string
building. `tupleToCoordinate(col, row)` is close but takes column first and has no `$` support.

### 5. Formula plus format in one write

`setCellFormula`, `setCellArrayFormula` and `setCellRichText` still exist in
`src/worksheet/worksheet.ts` but were removed from the public surface in the "one way per task"
trim (commit `8ce7059`) as redundant with `setCell` + `setFormula`. They are now dead code and
should be deleted, or the decision reversed.

The PR argues for reversal in practice: every formula cell is written as
`setFormula(cellAt(...), text)` and then formatted, which is the exact body of the removed helper.
With `registerCellStyle` from item 3 the ergonomic path becomes
`setCell(ws, r, c, formulaValue, styleId)`, which needs a `formula(text)` value constructor in
`/cell` rather than a mutating `setFormula(cell, text)`.

Recommendation: add `makeFormula(text, { cachedValue? }): FormulaValue` in `/cell`, keep
`setFormula` as the mutate-in-place form, delete the three dead worksheet helpers.

### 6. Deterministic output

`createZipWriter` constructs `new ZipDeflate(path)` and never sets `mtime`, so fflate stamps
`Date.now()` into every local header and central directory entry. Two renders of the same payload
differ in bytes, which blocks golden-file tests and content-addressed caching. `SaveOptions` has a
single field, `compressionLevel`, documented as "Reserved" and not wired through.

Proposal: `SaveOptions.mtime?: Date` (or `timestamp`), applied to every entry via the
`ZipDeflate.mtime` / `ZipPassThrough.mtime` field before `zip.add`. The streaming writer should
take the same option. Wire or remove `compressionLevel` while touching the type.

With this in place, `workbookToBytes(wb, { mtime })` plus the caller-controlled core properties the
PR already sets gives byte-identical output for identical input.

### 7. Formulas without cached values

None of the renderer's formulas carry a `cachedValue`, and the workbook does not set
`fullCalcOnLoad`. Excel, Sheets and LibreOffice compute a formula with no `<v>` on open, so the
report is correct where the author tested it. Viewers that do not calculate (Quick Look, Outlook
and SharePoint previews, most thumbnailers) show those cells empty, and the reconciliation cell is
one of them.

The library already offers both fixes: `setFormula(cell, text, { cachedValue })` and
`setFullCalcOnLoad(wb, true)`. Neither is mentioned in the formula recipe beyond a one-line comment
about the cached value.

Proposal: a recipe section "Formulas in generated files" that explains the three viewer classes,
recommends `setFullCalcOnLoad(wb, true)` for any generated workbook with formulas, and shows how to
supply a cached value when the producer can compute it (the renderer can, for every formula it
writes except the `SUMIFS` share).

### 8. Testing a generated workbook

The PR reimplements three helpers that any renderer test needs. The API trim intentionally removed
the JSON and matrix exports, so this is a documentation item, not a request to bring them back:

- a recipe "Assert on a generated workbook" showing `loadWorkbook(fromArrayBuffer(bytes))`,
  `getSheet` narrowing, `iterCells` plus `getFormulaText`, `getRangeValues`, `getAutoFilter`
- a note that `fromArrayBuffer` accepts a `Uint8Array` directly. Pre-1.0 a rename to `fromBytes`
  is allowed by the one-way rule if the current name is judged misleading; otherwise a docstring
  line is enough
- a note that `addWorksheet` validates the sheet title (length, reserved characters, `History`), so
  consumers do not need a 31-character test of their own

### 9. Excel Tables for tidy ranges

The "Leverage by language" sheet is designed as a tidy, filterable, pivotable range and uses an
autofilter with hand-drawn header styling and per-cell borders. An Excel Table (`addExcelTable`
with a `style`) gives the same filter plus banded rows, structured references and a stable name
for pivots, in one call and with no per-cell border work.

The recipe "tables-with-filter" exists. Two things would make it land for this use case:

- the cheatsheet has no row for "turn a data range into an Excel Table"
- `addExcelTable` does not check that the header cells in `ref` match `columns`, and Excel repairs
  the file when they disagree. A boundary check here follows the "validate at boundaries" rule
  and would have made the migration safe to try.

### 10. Smaller items

- **Two ways to freeze panes.** `setFreezePanes(ws, 'A5')` and `freezePanes(ws, rows, cols)` are
  both exported. The PR used the A1 form; the numeric form matches how it tracks rows. Pick one or
  document when each applies.
- **Number-format helpers exist and were not found.** `setCellAsNumber(wb, c)` produces `#,##0`
  and `setCellAsPercent(wb, c, 1)` produces `0.0%`, the two codes the PR defines as constants.
  Cheatsheet discoverability.
- **Input cells.** The empty, editable multiplier column would benefit from
  `applyBuiltinStyle(wb, c, 'Input')` and an `addDataValidation` decimal rule. Both exist; a
  "user-editable input column" recipe would surface them.
- **`autofitColumns`** exists and would replace the hard-coded width arrays for the text-heavy
  sheets. Cheatsheet row.

## Priority summary

| # | Item                                              | Kind                  | Breaking | Effort |
| - | ------------------------------------------------- | --------------------- | -------- | ------ |
| 1 | `ensureCell` + required `value` on `setCell`      | API + docs            | Yes      | Small  |
| 3 | `registerCellStyle` returning a `styleId`         | API                   | No       | Small  |
| 6 | `SaveOptions.mtime` for deterministic bytes       | API                   | No       | Small  |
| 2 | `patchCellFont`, rebuild single-field setters     | API + docs            | No       | Small  |
| 4 | Numeric bounds on range helpers, `cellRef` helper | API                   | No       | Medium |
| 5 | `makeFormula` value constructor, delete dead code | API + cleanup         | No       | Small  |
| 7 | Formula recipe: cached values, `fullCalcOnLoad`   | Docs                  | No       | Small  |
| 8 | Recipe: asserting on a generated workbook         | Docs                  | No       | Small  |
| 9 | Table header check, cheatsheet row                | Validation + docs     | No       | Small  |
| 10| Freeze-panes duplication, cheatsheet rows         | Docs + API tidy       | Maybe    | Small  |

Items 1, 3 and 6 together would have removed the `put` / `cellAt` split, both per-row style loops,
the four-call header helper and the content-based determinism test from the PR.
