# Migrating from SheetJS to @office-kit/xlsx

SheetJS reads and writes a dozen spreadsheet formats through one worksheet
object and a set of conversion helpers. `@office-kit/xlsx` does one format,
`.xlsx` / `.xlsm`, through a typed model of the OOXML parts, and exposes it as
free functions behind per-area subpath imports. The translation is mechanical
for the reading and writing paths, and the shape of the code changes in three
places worth knowing before you start.

**There is no root import.** Every export lives behind the subpath that owns it:
`@office-kit/xlsx/io`, `/workbook`, `/worksheet`, `/cell`, `/styles`. Replace
`import * as XLSX from 'xlsx'` with named imports from the two or three subpaths
you use.

**Reading is async.** `loadWorkbook` takes an `XlsxSource` (`fromFile`,
`fromBuffer`, `fromBlob`, `fromResponse`, `fromStream`) and returns a promise,
so the parse can stream and the same call works in Node and the browser.

**Cells are a discriminated union, not a `{ t, v, w }` record.** A cell's
`value` is a `number | string | boolean | Date`, a formula, rich text, an error
or `null`. There is no `w` field holding a pre-rendered string: the text is
computed on demand from the cell's number format by `getCellDisplayText`.

## Reading a file

```js
// SheetJS
const XLSX = require('xlsx');
const wb = XLSX.read(fs.readFileSync('input.xlsx'));
```

```ts
// @office-kit/xlsx
import { loadWorkbook } from '@office-kit/xlsx/io';
import { fromFile } from '@office-kit/xlsx/node';

const wb = await loadWorkbook(fromFile('input.xlsx'));
```

`loadWorkbook` validates as it parses and throws an `OpenXmlError` subclass on
malformed input rather than returning a partial workbook. It also applies
`decompressionLimits` by default, which bounds what an adversarial archive can
cost you. Leave those on for anything you did not produce yourself.

## Writing a file

```js
// SheetJS
XLSX.writeFile(wb, 'output.xlsx');
```

```ts
// @office-kit/xlsx
import { saveWorkbook } from '@office-kit/xlsx/io';
import { toFile } from '@office-kit/xlsx/node';

await saveWorkbook(wb, toFile('output.xlsx'));
```

`workbookToBytes(wb)` returns the bytes instead, for an HTTP response or a
`Blob`.

## API map

| SheetJS                               | @office-kit/xlsx                                          |
| ------------------------------------- | --------------------------------------------------------- |
| `XLSX.read(data)`                     | `await loadWorkbook(fromBuffer(data))` (`/io`)            |
| `XLSX.readFile(path)`                 | `await loadWorkbook(fromFile(path))` (`/io` + `/node`)    |
| `XLSX.write(wb)`                      | `await workbookToBytes(wb)` (`/io`)                       |
| `XLSX.writeFile(wb, path)`            | `await saveWorkbook(wb, toFile(path))` (`/io` + `/node`)  |
| `wb.SheetNames`                       | `sheetNames(wb)` (`/workbook`)                            |
| `wb.Sheets[name]`                     | `getSheet(wb, name)` (`/workbook`)                        |
| `XLSX.utils.book_new()`               | `createWorkbook()` (`/workbook`)                          |
| `XLSX.utils.book_append_sheet(wb, ws, name)` | `addWorksheet(wb, name)` (`/workbook`)             |
| `ws['A1']`                            | `ensureCellByCoord(ws, 'A1')` (`/worksheet`)              |
| `ws['A1'].v`                          | `getCellByCoord(ws, 'A1')?.value` (`/worksheet`)          |
| `ws['A1'].w`                          | `getCellDisplayText(wb, cell)` (`/styles`)                |
| `ws['A1'].f`                          | `getFormulaText(cell)` (`/cell`)                          |
| `ws['A1'].z`                          | `getCellNumberFormat(wb, cell)` (`/styles`)               |
| `XLSX.utils.sheet_to_json(ws, { header: 1 })` | `[...iterValues(ws)]` (`/worksheet`)              |
| `XLSX.utils.sheet_to_json(ws)`        | `iterValues(ws)` plus your own header row handling        |
| `XLSX.utils.aoa_to_sheet(rows)`       | `appendRows(ws, rows)` (`/worksheet`)                     |
| `XLSX.utils.encode_cell({ r, c })`    | `tupleToCoordinate(col, row)` (`/utils`)                  |
| `XLSX.utils.decode_cell('B3')`        | `coordinateToTuple('B3')` (`/utils`)                      |
| `XLSX.utils.decode_range('A1:C9')`    | `rangeBoundaries('A1:C9')` (`/utils`)                     |
| `ws['!merges']`                       | `getMergedCells(ws)` (`/worksheet`)                       |
| `ws['!cols']`                         | `setColumnWidth(ws, col, width)` (`/worksheet`)            |
| `ws['!freeze']` / `XLSX.utils.*`      | `setFreezePanes(ws, 'B2')` (`/worksheet`)                  |
| `XLSX.stream.to_csv(ws)`              | `iterValues(ws)` plus your own CSV writer                 |

Indexes are 1-based in `@office-kit/xlsx`, matching the A1 references in the
file. SheetJS's `{ r, c }` are 0-based, so add one when you port arithmetic.

## `sheet_to_json` and the `raw` option

`sheet_to_json(ws, { header: 1 })` gives you an array of row arrays. The
equivalent is `iterValues`, which is a generator, so you can consume it row by
row or spread it:

```ts
import { iterValues } from '@office-kit/xlsx/worksheet';

for (const row of iterValues(ws)) {
  // row is CellValue[]; a gap in the row is null
}
```

`iterValues` yields values, which is `raw: true`. For `raw: false`, where
SheetJS hands you the formatted string, walk the cells and ask for the text:

```ts
import { getCellDisplayText } from '@office-kit/xlsx/styles';
import { iterRows } from '@office-kit/xlsx/worksheet';

const text = [...iterRows(ws)].map((row) =>
  row.map((cell) => (cell === undefined ? '' : getCellDisplayText(wb, cell))),
);
```

`getCellDisplayText` needs the workbook because the number format lives in the
workbook stylesheet, not on the cell. It renders `0.5` under `0.0%` as `50.0%`,
a serial under `yyyy-mm-dd` as `2024-03-14`, and `1.1 + 2.2` under `General` as
`3.3`. Its docstring lists the format codes it covers and the ones it falls back
on.

## Options with no equivalent

**`cellDates`**: not applicable. SheetJS can turn date-formatted cells into
`Date` objects at parse time. `@office-kit/xlsx` does not, on purpose: a load
option that changes the type of a cell value makes every downstream `typeof`
check conditional on how the file was opened. The value model stays as the file
stores it, a number, and you convert where you need a date:

```ts
import { getCellDate } from '@office-kit/xlsx/styles';

const due = getCellDate(wb, cell); // Date, or undefined if the cell is not a date
```

`getCellDate` resolves the cell's number format, checks that it names a
calendar date, and converts the serial under the workbook epoch (`wb.date1904`).
It returns `undefined` for a number the format says is a count, for a
time-of-day format such as `h:mm`, which names a moment inside a day but not
which day, and for an elapsed-time format such as `[h]:mm:ss`, which measures a
span rather than naming a moment.

**`codepage`**: not applicable. A code page is a legacy-format concern. The
parts inside an xlsx package are UTF-8 XML, so the text arrives as Unicode with
nothing to guess and nothing to override. `codepage` matters for `.xls` and
`.csv`, which this library does not read.

**`cellStyles`**: always on. Styles are part of the model: the cell carries a
`styleId` into `wb.styles.cellXfs`, and the `@office-kit/xlsx/styles` bridge
(`getCellFont`, `getCellFill`, `getCellNumberFormat`, `setCellFont`, ...) reads
and writes through it.

**`sheetRows`**: use the iteration bounds instead. `iterRows(ws, { maxRow })`
stops there, and for a file too large to hold in memory `loadWorkbookStream`
from `@office-kit/xlsx/streaming` stops reading the bytes once it crosses
`maxRow`.

**`bookVBA`**: on by default. `xl/vbaProject.bin` round-trips byte-identical and
the writer promotes the package to `.xlsm` when it is present.

## Formats out of scope

`@office-kit/xlsx` reads and writes `.xlsx` and `.xlsm` only. For anything else,
stay on SheetJS or reach for a format-specific library:

| Format                     | Where to go                                                |
| -------------------------- | ---------------------------------------------------------- |
| `.csv` / `.txt`            | Any CSV library, or `iterValues` plus a writer of your own |
| `.xls` (BIFF)              | SheetJS                                                    |
| `.xlsb` (binary workbook)  | SheetJS                                                    |
| `.ods` / `.fods`           | SheetJS                                                    |
| `.numbers`                 | SheetJS                                                    |
| Encrypted xlsx             | Decrypt with [`msoffcrypto-tool`][msoffcrypto] first       |

[msoffcrypto]: https://github.com/nolze/msoffcrypto-tool

Passing a `.xls` or an encrypted xlsx to `loadWorkbook` throws
`OpenXmlUnsupportedFormatError`, whose `format` field is `'legacy-xls'` or
`'encrypted-xlsx'`, so an upload handler can route the file to SheetJS or ask
for the password without matching on the message.

The narrower surface is the trade. One format means the model can be typed to
the OOXML schema rather than to a lowest common denominator, so charts, pivot
pass-through, tables, data validation and conditional formatting are reachable
with types that describe what Excel actually accepts.

## What you gain in the port

- **Types that match the file.** Cell values are a discriminated union, so a
  formula cell cannot be read as a number by accident, and adding a variant is
  a compile error at every `switch`.
- **Loud failures.** Malformed input throws an `OpenXmlError` subclass with a
  `cause` chain instead of yielding a workbook with holes in it.
- **Bounded cost on untrusted input.** Per-entry, per-archive and compression
  ratio caps are on by default; see `SECURITY.md`.
- **Streaming both ways.** `loadWorkbookStream` reads a sheet larger than
  memory; `createWriteOnlyWorkbook` writes one.

## Further reading

- `README.md` for the feature matrix, subpath entries and bundle budgets.
- [Recipes](https://baseballyama.github.io/@office-kit/xlsx/docs/recipes),
  including "Read a workbook somebody else produced".
- `docs/migrate-from-openpyxl.md` if you are also porting a Python producer.
- `SECURITY.md` for `decompressionLimits` and the threat model.
