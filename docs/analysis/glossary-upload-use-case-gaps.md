# Use-case analysis: glossary upload parser (autopilot `backend-file-conversion`)

Analysis of how `services/in-and-out/backend-file-conversion` in `lokalise/autopilot` uses
`@lokalise/xlsx@0.20.3` (a fork of SheetJS Community 0.20.3), and what `@office-kit/xlsx` needs
before that service can swap libraries without changing what users see. Measured against this
repository at `32e67d0` (`0.17.0`). No code changes are proposed here; this is input for scoping.

## The use case

One file, `src/modules/upload/conversion/services/GlossaryParser.ts`, turns an uploaded glossary
into terms:

```ts
workbook = XLSX.read(await rawContent.arrayBuffer(), { type: 'array', codepage: 65001 })
const sheet = workbook.Sheets[workbook.SheetNames[0]]
const content: string[][] = XLSX.utils.sheet_to_json(sheet, { header: 1, raw: false })
```

`content` then goes to a header mapper that resolves column indexes by name, and a row mapper that
trims every value, coerces `y` / `yes` / `1` / `true` to booleans, and splits comma lists. Nothing
downstream ever sees a number, a `Date` or a cell style: the whole pipeline is text keyed by column
index.

### Library surface actually used

| Symbol                                        | Purpose                                      |
| --------------------------------------------- | -------------------------------------------- |
| `XLSX.read(bytes, { type, codepage })`        | parse the uploaded bytes                     |
| `XLSX.WorkBook`, `.SheetNames`, `.Sheets`     | reach the first sheet                        |
| `XLSX.utils.sheet_to_json(sheet, { header: 1, raw: false })` | whole sheet as `string[][]` of display text |

That is the entire dependency: five symbols, read-only, no writing, no styling, no streaming.

### Facts that shape the migration

- Uploads are restricted to two extensions. `GlossaryUploadService.validateFileExtension` accepts
  `.csv` and `.xlsx` only, and the resolved `fileType: 'CSV' | 'XLSX'` rides on the conversion job
  payload. The parser ignores it today and lets SheetJS sniff the bytes.
- Both formats go through `XLSX.read`. The CSV fixture in the repo is semicolon-delimited and the
  tests feed comma-delimited strings, so the service silently depends on SheetJS's delimiter
  detection. `codepage: 65001` only ever mattered for that branch; xlsx stores XML, which is
  Unicode by definition.
- Failures collapse into a single outcome. A throw from `read` is caught, logged at `warn`, and
  returned as `{ errors: [{ reason: 'SERVER_ERROR' }] }`, which the job turns into an unrecoverable
  failure and an error event.
- The service logs a warning when a parse takes 50 ms or more. Tests cover a 5000-row file.
- The xlsx test fixture contains a drawing part. It loads through `loadWorkbook` without complaint,
  so container fidelity is not at issue here.

## The shape of the migration

1. **CSV leaves the xlsx library.** `.csv` is out of scope by design (`README.md` says so, and
   `CLAUDE.md` routes csv to SheetJS). The job payload already carries the type, so the parser
   branches on `fileType` and the CSV branch gets a real CSV parser with delimiter detection. This
   needs nothing from `@office-kit/xlsx`.
2. **The xlsx branch needs `Blob` to `string[][]`.** `fromBlob`, `loadWorkbook`, the first sheet,
   rows of display text.

Everything below is what stands between step 2 and a behaviour-preserving swap.

## Measurements

Same process, same bytes, Node 24.18 on Windows, files written by this repository's writer with six
text columns, best of three runs. `office-kit` is `loadWorkbook` plus `iterValues` mapped through
`cellValueAsString`; `sheetjs` is `read` plus `sheet_to_json({ header: 1, raw: false })`.

| Workload                                            | office-kit | sheetjs | ratio  |
| --------------------------------------------------- | ---------- | ------- | ------ |
| 5 000 rows (172 KiB)                                | 224 ms     | 75 ms   | 3.0x   |
| 50 000 rows (1 691 KiB)                             | 2 370 ms   | 886 ms  | 2.7x   |
| retained heap after parse, 50 000 rows              | 52 MB      | 211 MB  | 0.25x  |
| peak RSS over a 103 MB idle baseline, 50 000 rows   | 673 MB     | 423 MB  | 1.6x   |

The wall-clock rows were taken while other work ran on the machine, so read the absolute figures as
an upper bound; the ratio held between 2.2x and 3.0x across every run, on an idle machine and a busy
one. `loadWorkbook` is 98% of the office-kit figure: at 50 000 rows the parse is 2 324 ms and
turning the model into `string[][]` is 46 ms. The streaming reader parses the same file in about the
same time as the full load, so the cost sits in the XML read, not in model construction.

## Findings, in priority order

### 1. A single unmodelled cell aborts the whole file

Two shapes that Excel writes today make `loadWorkbook` throw, verified by patching the service's own
fixture:

```
#SPILL! in <c t="e">                   OpenXmlSchemaError: worksheet: unknown error code "#SPILL!" in <c t="e">
#CALC!  in <c t="e">                   OpenXmlSchemaError: worksheet: unknown error code "#CALC!" in <c t="e">
<c t="d"><v>2024-03-14T00:00:00</v>    OpenXmlSchemaError: worksheet: unknown cell type t="d"
```

`ERROR_CODES` in `src/utils/inference.ts` holds the eight pre-2018 tokens. The dynamic-array and
data-type errors Excel 365 emits (`#SPILL!`, `#CALC!`, `#FIELD!`, `#BLOCKED!`, `#CONNECT!`,
`#BUSY!`, `#UNKNOWN!`, `#PYTHON!`, `#EXTERNAL!`) are absent. `t="d"` is an ISO 29500 cell type
carrying an ISO 8601 date in `<v>`; it is not in the transitional XSD this repository validates
against, but producers emit it and SheetJS reads it.

The streaming reader disagrees with the DOM reader on the same input: `src/streaming/read-only.ts`
returns `null` for a token outside the set, so one path drops the cell and the other aborts the file.
Both are wrong, in opposite directions.

For a glossary importer the loss is total and the message is unusable: 5 000 good rows fail because
one cell somewhere holds a spilled formula, and the user is told `SERVER_ERROR`. This is the finding
that can turn a working upload into a broken one, which is why it leads.

Fix: tolerate both on read. Add the Excel 365 tokens to `ERROR_CODES` and `ExcelErrorCode`, and
handle `t="d"` in the cell-type switch. Keep throwing for input that genuinely cannot be parsed.
Widening `ExcelErrorCode` is additive, but it does break an exhaustive `switch` over the union in
consumer code, so it belongs in a changeset.

### 2. There is no display text, and date cells read as serial numbers

`raw: false` asks SheetJS for what Excel shows. `@office-kit/xlsx` has no equivalent:
`cellValueAsString` is a JavaScript coercion that never looks at the cell's number format.

| Cell                             | Excel / sheetjs `raw: false` | `cellValueAsString`  |
| -------------------------------- | ---------------------------- | -------------------- |
| `0.5` with `0.0%`                | `50.0%`                      | `0.5`                |
| `1234567.891` with `#,##0.00`    | `1,234,567.89`               | `1234567.891`        |
| `1.1 + 2.2` with `General`       | `3.3`                        | `3.3000000000000003` |
| serial `45365` with `yyyy-mm-dd` | `2024-03-14`                 | `45365`              |
| `true`                           | `TRUE`                       | `true`               |

The date row produces nonsense rather than a cosmetic difference: a glossary cell that Excel typed
as a date arrives as `45365`, and the mapper stores that as the term. The recovery path exists but
is neither composed nor documented: read `getCellNumberFormat` from `/styles`, test it with
`isDateFormat`, then call `excelToDate` from `/utils` with the epoch from `wb.date1904`.
`cellValueAsDate` deliberately does not do this, and says so in its docstring.

The `General` row matters too. SheetJS renders numbers through a significant-digit pass, so float
noise never reaches the consumer; `String(v)` hands it over.

Fix: a format-aware renderer, for example `formatCellText(wb, cell): string` in `/styles`, covering
`General` (with significant-digit trimming), the built-in format ids, the common custom numeric
codes (`#,##0`, `0.00`, percent, currency) and date / time codes, with Excel's own rendering of
booleans and error tokens. Outside that set, degrade to the plain coercion and document the limit
rather than guessing.

On the one-way rule: this is not a parallel path to `cellValueAsString`. That function takes a
`CellValue` and answers "what is this value in JavaScript terms"; the renderer needs the workbook's
stylesheet and answers "what does Excel print in this cell". Different inputs, different answers.
The docs have to say which one a reader wants, because the wrong pick fails silently.

A `LoadOptions.cellDates` flag is the other option and is worse: it makes the value model depend on
a load flag, so every downstream `typeof` check becomes conditional. The renderer plus one
documented date helper covers the same ground without that.

### 3. Formatting alone decides where the data ends

Append an empty `<row r="5"/>` and a style-only cell (`<row r="6"><c r="A6" s="0"/></row>`) to the
service's four-row fixture and `getMaxRow` returns 6, `getDataExtent` reports `maxRow: 6`, and
`iterValues` yields two trailing rows of `null`. Nothing was typed into either row.

`getMaxRow`, `getMaxCol` and `getDataExtent` count any cell *object* in `ws.rows`, and a cell that
carries only a style is an object. A sheet where somebody formatted 200 rows and typed data in 4
iterates 196 empty rows, and a caller mapping rows to records gets 196 phantom records. For this
consumer they arrive as `MISSING_REQUIRED_COLUMN_WARNING` lines in a user's import report.

The two readings are both legitimate and Excel keeps them apart. Its used range, and the
`<dimension>` it writes, include cells that hold only formatting; what a person means by "my data"
is the cells that hold values. openpyxl's inflated `max_row` is the well-known version of this trap.

Fix: make the distinction explicit rather than picking one silently. A public value-based extent next
to `getDataExtent`, an option on `IterRowsOptions` selecting which extent bounds the iteration with
today's behaviour as the default, and docstrings that say which extent matches Excel's used range,
which one answers "where is the data", and what the padding contract is: rows are padded to the
extent width, unpopulated positions are `null`, and a position in a yielded row maps to a fixed
column.

Not worth doing: matching SheetJS, which yields `[]` for a blank row and truncates each row at its
last populated cell. Rectangular rows with `null` holes are the better model, because the position
of a value in the yielded row is what a header-driven reader resolves columns by, and ECMA-376
treats an absent row, an empty `<row/>` and a formatted-but-valueless row as the same thing. The
consumer's own mapper carries a comment about working around SheetJS truncating row tails, which
says which side of this is the footgun. An option that drops blank rows from the middle of the
sequence is worse still, because it renumbers rows silently; a caller who wants it can write
`filter(row => row.some(v => v !== null))`, which the `iterRows` docstring already teaches.

One hole semantic does need writing down. `iterValues` returns `null` at unpopulated positions, and
the mapper reads `row[index] ?? ''`, so a `null` survives into `getCoercedValue` where `value.trim()`
throws on it and the row is reported as `FAILED_TO_PARSE_LINE_ERROR`. Mapping through
`cellValueAsString` avoids that, and the read recipe has to say so.

### 4. Read throughput is 2.4x to 3x SheetJS

The numbers are in the measurements table. `src/worksheet/reader.ts` opens with the reason: the
`sheetData` reader is "Stage 1: DOM-based", with SAX `iterparse` planned. Building an `XmlNode` tree
for every `<c>` before producing a `Cell` is where both the wall-clock gap and the 1.6x peak RSS
come from. The retained model is four times smaller than SheetJS's, so this is a transient-cost
problem, not a footprint problem.

Fix: move `sheetData` to `iterParse`, which is already the stated plan, and add a read-direction
case to `tests/perf/` so the gain is defended by the gate rather than by a one-off measurement.

Worth telling the consumer either way: their parse-time warning fires at 50 ms and a 5 000-row file
costs a couple of hundred milliseconds, so the threshold needs revisiting as part of the swap. It is
a log line, not a failure.

### 5. Error taxonomy for "the user uploaded the wrong thing"

What the loader throws for input a user can actually produce:

| Input                               | Error                                                                               |
| ----------------------------------- | ----------------------------------------------------------------------------------- |
| CSV bytes, plain text, random bytes | `OpenXmlIoError: openZip: archive is not a valid zip`                               |
| empty file                          | `OpenXmlIoError: openZip: archive is shorter than the minimum EOCD size (22 bytes)` |
| valid zip, damaged part             | `OpenXmlSchemaError: parseXml: failed to parse XML payload`                          |

Every path is a typed `OpenXmlError`, which is the hard part and is already done. What is missing is
the contract a consumer needs to split "bad upload, tell the user, do not retry" from "our bug,
retry and page someone". Today that decision would be made by matching message strings.

This matters more than it looks, because the extension check upstream only reads the filename. A CSV
saved as `.xlsx`, or a real `.xls` renamed, reaches the loader and fails as "not a valid zip", which
is true and unhelpful.

Fix, in order of value:

- Document which error classes mean "the input was not a valid xlsx" and which mean "the library
  failed", as a table next to `loadWorkbook`. This is the actual blocker.
- In the not-a-zip path, name what the bytes look like from their magic number: `D0 CF 11 E0` is a
  legacy `.xls`, `%PDF` is a PDF, a BOM or all-printable content is text or CSV. One sniff turns the
  message into something a service can put in front of a user.

### 6. No workload bound for untrusted uploads

`LoadOptions` exposes `decompressionLimits` only, defaulting to 512 MiB per entry, 1 GiB total and a
ratio of 1000. Those are generous by design and nowhere near what a glossary needs, and they bound
inflated bytes rather than the thing that decides heap, which is cell count. A 2 MB upload that
inflates to a few hundred MB of `sheetData` is within the defaults, and a model gets built for every
cell before the caller can look at it.

Fix: `LoadOptions` gains a cell or row budget (`maxCells`, `maxRows`) that throws a typed error when
exceeded, and the docs grow an "ingestion profile" showing tighter limits for services that accept
files from strangers. The library already takes that posture for zip bombs; cell count is the same
argument one layer up.

### 7. Reaching the first sheet is not part of the public surface

`workbook.SheetNames[0]` becomes `sheetNames(wb)[0]` followed by `getSheet(wb, name)`: two lookups
and two non-null assertions in a codebase that runs `noUncheckedIndexedAccess`.
`getSheetByIndex(wb, idx)` already exists in `src/workbook/workbook.ts` and is not exported from
`src/workbook/index.ts`. `hasSheet` and `getDataExtentRef` are in the same position.

Fix: export `getSheetByIndex`, or delete it. Same call on the other two. This is the `ensureCell`
lesson from the leverage-report analysis repeating: the helper a consumer needs is written, used
internally, and invisible.

### 8. ISO strict workbooks are rejected at the first step

Excel's "Strict Open XML Spreadsheet" save option produces a file with the `.xlsx` extension and
`purl.oclc.org` namespaces. It passes the service's extension check and fails here:

```
OpenXmlSchemaError: loadWorkbook: root rels missing officeDocument relationship
```

SheetJS reads these files. Frequency is low, since nobody picks that entry by accident, but the
failure is total and the message points at the wrong thing.

Fix: accept both namespace families on read through an alias table for the relationship types and
element namespaces. Medium effort. If that is not worth scheduling, detect the strict namespace and
throw a message that names it, which costs one comparison.

### 9. Documentation: a read-path recipe and a SheetJS migration map

The recipes cover producing workbooks. This consumer only reads, and every finding above has a
documentation half:

- A recipe for reading a user-supplied workbook: `fromBlob`, first sheet, `isWorksheetEmpty`, rows
  to text, blank-row handling, date-formatted cells, and the error table from finding 5.
- `docs/migrate-from-sheetjs.md`, next to the existing openpyxl map: `read` to `loadWorkbook`,
  `SheetNames` to `sheetNames`, `Sheets[name]` to `getSheet`, `sheet_to_json({ header: 1 })` to
  `iterValues`, `raw: false` to the renderer from finding 2, `codepage` not applicable and why,
  `cellDates` not applicable and what to do instead, and `.csv` / `.xls` / `.ods` out of scope with
  a pointer. The README already positions the library against SheetJS, and a SheetJS fork sitting in
  Lokalise's own dependency tree makes this the most likely inbound migration there is.

## Priority summary

| # | Item                                                         | Kind             | Breaking                | Effort |
| - | ------------------------------------------------------------ | ---------------- | ----------------------- | ------ |
| 1 | Excel 365 error tokens, `t="d"` cells                        | Fix              | Additive union widening | Small  |
| 2 | Format-aware `formatCellText`, date-cell reads               | API + docs       | No                      | Large  |
| 3 | Value-based extent, extent option on row iteration, padding contract documented | API + docs | No | Small |
| 5 | Error-class contract documented, format sniff in the message  | Docs + small API | No                      | Small  |
| 7 | Export `getSheetByIndex` (and decide on `hasSheet`)          | API              | No                      | Small  |
| 4 | `sheetData` to SAX, read-direction perf gate                 | Performance      | No                      | Medium |
| 6 | `maxCells` / `maxRows` load budget, ingestion profile docs    | Security + docs  | No                      | Medium |
| 9 | Read recipe, `migrate-from-sheetjs.md`                       | Docs             | No                      | Small  |
| 8 | ISO strict namespaces on read                                | API              | No                      | Medium |

Items 1, 2 and 3 are the migration blockers: without them the swap either fails on files that work
today, or changes what the user sees in their import report. Items 5 and 7 are half a day and remove
the awkward parts of the port. Item 4 is invisible to correctness and should be judged on the
library's own roadmap, not on this consumer, whose files are small.

## What this consumer does not need

Listed so the items above are not read as a larger mandate: no writing, no styles, no charts, no
streaming (the whole sheet has to be in memory anyway, and 50 000 rows fit), no CSV inside this
library, no codepage handling, no formula evaluation, and no fidelity guarantees for drawings or
pivots. The read path is the entire surface.
