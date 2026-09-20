// Cell shapes a current Excel writes that the reader used to refuse: the error
// tokens added since 2018, and the `t="d"` ISO-8601 date of ISO 29500 strict.
// One such cell anywhere in a part aborted the whole load, so a sheet of good
// rows was lost to a single spilled formula.

import { unzipSync, zipSync } from 'fflate';
import { describe, expect, it } from 'vitest';
import { type ExcelErrorCode, isErrorValue, makeFormula } from '../../src/cell/cell.js';
import { loadWorkbook } from '../../src/io/load.js';
import { fromBuffer } from '../../src/io/node.js';
import { workbookToBytes } from '../../src/io/save.js';
import { OpenXmlSchemaError } from '../../src/utils/exceptions.js';
import { ERROR_CODES } from '../../src/utils/inference.js';
import { addWorksheet, createWorkbook, getSheet, type Workbook } from '../../src/workbook/workbook.js';
import { parseWorksheetXml } from '../../src/worksheet/reader.js';
import { getCell, setCell } from '../../src/worksheet/worksheet.js';

const MAIN_NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const SHEET_PART = 'xl/worksheets/sheet1.xml';

const sheet = (body: string): string =>
  `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
  `<worksheet xmlns="${MAIN_NS}"><sheetData>${body}</sheetData></worksheet>`;

const readSheet = (body: string) => parseWorksheetXml(sheet(body), 'Data', { sharedStrings: [] });

/**
 * A workbook saved by the library with the `<c>` of its one cell swapped out.
 * Going through the real save path keeps content types, rels and the workbook
 * part exactly as a legitimate file has them, so the load reaches the cell.
 */
const savedWithCell = async (cellXml: string, date1904 = false): Promise<Uint8Array> => {
  const wb = createWorkbook({ date1904 });
  setCell(addWorksheet(wb, 'Data'), 1, 1, 42);
  const entries = unzipSync(await workbookToBytes(wb));
  const part = entries[SHEET_PART];
  if (part === undefined) throw new Error(`save produced no ${SHEET_PART}`);
  const patched = new TextDecoder().decode(part).replace('<c r="A1"><v>42</v></c>', cellXml);
  entries[SHEET_PART] = new TextEncoder().encode(patched);
  return zipSync(entries);
};

/** Load a patched single-cell workbook and hand back its one worksheet. */
const loadPatched = async (cellXml: string) => {
  const wb = await loadWorkbook(fromBuffer(await savedWithCell(cellXml)));
  const ws = getSheet(wb, 'Data');
  if (ws === undefined) throw new Error('load produced no worksheet');
  return { wb, ws };
};

const savedSheetText = async (wb: Workbook): Promise<string> => {
  const part = unzipSync(await workbookToBytes(wb))[SHEET_PART];
  if (part === undefined) throw new Error(`save produced no ${SHEET_PART}`);
  return new TextDecoder().decode(part);
};

const MODERN_CODES = [
  '#SPILL!',
  '#CALC!',
  '#FIELD!',
  '#BLOCKED!',
  '#CONNECT!',
  '#BUSY!',
  '#UNKNOWN!',
  '#PYTHON!',
  '#EXTERNAL!',
] as const;

describe('error tokens Excel added after 2018', () => {
  it('reads each one as an error cell', () => {
    for (const code of MODERN_CODES) {
      const ws = readSheet(`<row r="1"><c r="A1" t="e"><v>${code}</v></c></row>`);
      expect(getCell(ws, 1, 1)?.value).toEqual({ kind: 'error', code });
    }
  });

  it('carries #SPILL! through loadWorkbook', async () => {
    const { ws } = await loadPatched('<c r="A1" t="e"><v>#SPILL!</v></c>');
    expect(getCell(ws, 1, 1)?.value).toEqual({ kind: 'error', code: '#SPILL!' });
  });

  it('lists them in ERROR_CODES, so a caller can name one on a write too', () => {
    for (const code of MODERN_CODES) expect(ERROR_CODES.has(code)).toBe(true);
  });

  it('writes them back as t="e" verbatim', async () => {
    const { wb } = await loadPatched('<c r="A1" t="e"><v>#CALC!</v></c>');
    expect(await savedSheetText(wb)).toContain('t="e"><v>#CALC!</v>');
  });
});

describe('an error token the library does not model', () => {
  it('keeps the token rather than dropping the cell', () => {
    const ws = readSheet('<row r="1"><c r="A1" t="e"><v>#NOTYET!</v></c></row>');
    const value = getCell(ws, 1, 1)?.value ?? null;
    expect(isErrorValue(value)).toBe(true);
    expect(value).toEqual({ kind: 'error', code: '#NOTYET!' });
  });

  it('survives a load then save', async () => {
    const { wb } = await loadPatched('<c r="A1" t="e"><v>#NOTYET!</v></c>');
    expect(await savedSheetText(wb)).toContain('t="e"><v>#NOTYET!</v>');
  });

  it('rejects a t="e" payload that is not an error token at all', () => {
    expect(() => readSheet('<row r="1"><c r="A1" t="e"><v>oops</v></c></row>')).toThrow(OpenXmlSchemaError);
    expect(() => readSheet('<row r="1"><c r="A1" t="e"><v>oops</v></c></row>')).toThrow(
      'worksheet: <v>oops</v> at Data!A1 is not an Excel error token',
    );
  });

  it('reads a blank value as an empty cell, as t="n" and t="b" do', () => {
    const ws = readSheet('<row r="1"><c r="A1" t="e"/><c r="B1" t="e"><v></v></c></row>');
    expect(getCell(ws, 1, 1)?.value).toBeNull();
    expect(getCell(ws, 1, 2)?.value).toBeNull();
  });

  it('carries the token through a formula cell rather than failing the save', async () => {
    const { wb, ws } = await loadPatched('<c r="A1" t="e"><f>A2</f><v>#NOTYET!</v></c>');
    expect(getCell(ws, 1, 1)?.value).toMatchObject({
      kind: 'formula',
      cachedValue: '#NOTYET!',
      cachedValueType: 'error',
    });
    expect(await savedSheetText(wb)).toContain('t="e"><f>A2</f><v>#NOTYET!</v>');
  });

  it('rejects a non-token cached formula error where the cell is, not on save', () => {
    expect(() => readSheet('<row r="1"><c r="A1" t="e"><f>A2</f><v>oops</v></c></row>')).toThrow(
      'worksheet: <v>oops</v> at Data!A1 is not an Excel error token',
    );
  });
});

// The token goes into `<v>` as it stands, so a `t="e"` payload outside the
// token shape would leave the saved part malformed, or close the `<v>` early
// and carry a cell of the file author's choosing into the sheet.
describe('a t="e" payload that is not shaped like a token', () => {
  const INJECTION = '#</v></c><c r="B1" t="s"><v>0</v></c>';

  it('is refused on read', () => {
    const escaped = INJECTION.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
    expect(() => readSheet(`<row r="1"><c r="A1" t="e"><v>${escaped}</v></c></row>`)).toThrow(
      OpenXmlSchemaError,
    );
    expect(() => readSheet('<row r="1"><c r="A1" t="e"><v>#A&amp;B</v></c></row>')).toThrow(
      'is not an Excel error token',
    );
  });

  it('is refused on write, naming the cell', async () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'Data');
    setCell(ws, 1, 1, { kind: 'error', code: INJECTION as ExcelErrorCode });
    await expect(savedSheetText(wb)).rejects.toThrow('worksheet: invalid error token at A1');
  });

  it('is refused as a cached formula result too', async () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'Data');
    setCell(ws, 1, 1, makeFormula('A2', { cachedValue: INJECTION, cachedValueType: 'error' }));
    await expect(savedSheetText(wb)).rejects.toThrow('worksheet: invalid cached formula error at A1');
  });
});

describe('t="d", the ISO 29500 strict date cell', () => {
  it('reads a naive datetime as UTC, the way every other Date in the model is read', () => {
    const ws = readSheet('<row r="1"><c r="A1" t="d"><v>2024-03-14T00:00:00</v></c></row>');
    expect(getCell(ws, 1, 1)?.value).toEqual(new Date(Date.UTC(2024, 2, 14)));
  });

  it('reads the date-only, fractional-second and offset forms', () => {
    const cases: ReadonlyArray<readonly [string, Date]> = [
      ['2024-03-14', new Date(Date.UTC(2024, 2, 14))],
      ['2024-03-14T12:30:45.500', new Date(Date.UTC(2024, 2, 14, 12, 30, 45, 500))],
      ['2024-03-14T12:30:45Z', new Date(Date.UTC(2024, 2, 14, 12, 30, 45))],
      ['2024-03-14T12:30:45+02:00', new Date(Date.UTC(2024, 2, 14, 10, 30, 45))],
      ['2024-03-14T12:30:45-05:30', new Date(Date.UTC(2024, 2, 14, 18, 0, 45))],
    ];
    for (const [text, expected] of cases) {
      const ws = readSheet(`<row r="1"><c r="A1" t="d"><v>${text}</v></c></row>`);
      expect(getCell(ws, 1, 1)?.value).toEqual(expected);
    }
  });

  it('reads a fraction longer than milliseconds, truncating it', () => {
    // Python's datetime.isoformat() prints six digits, so this is what every
    // openpyxl-written strict file carries.
    const ws = readSheet('<row r="1"><c r="A1" t="d"><v>2024-03-14T12:30:45.123456</v></c></row>');
    expect(getCell(ws, 1, 1)?.value).toEqual(new Date(Date.UTC(2024, 2, 14, 12, 30, 45, 123)));
  });

  it('reads a year before 100, which Date.UTC would otherwise read as 19xx', () => {
    const ws = readSheet('<row r="1"><c r="A1" t="d"><v>0099-01-01</v></c></row>');
    const value = getCell(ws, 1, 1)?.value;
    expect(value).toBeInstanceOf(Date);
    expect((value as Date).getUTCFullYear()).toBe(99);
  });

  it('reads the time-only and duration forms as a duration', () => {
    const ws = readSheet(
      '<row r="1"><c r="A1" t="d"><v>14:30:00</v></c><c r="B1" t="d"><v>PT4H30M</v></c></row>',
    );
    expect(getCell(ws, 1, 1)?.value).toEqual({ kind: 'duration', ms: 52_200_000 });
    expect(getCell(ws, 1, 2)?.value).toEqual({ kind: 'duration', ms: 16_200_000 });
  });

  it('reads a zoned time off the clock, since a fraction of a day cannot hold a zone', () => {
    const ws = readSheet(
      '<row r="1"><c r="A1" t="d"><v>14:30:00Z</v></c><c r="B1" t="d"><v>14:30:00+02:00</v></c></row>',
    );
    expect(getCell(ws, 1, 1)?.value).toEqual({ kind: 'duration', ms: 52_200_000 });
    expect(getCell(ws, 1, 2)?.value).toEqual({ kind: 'duration', ms: 52_200_000 });
  });

  it('reads 24:00:00, the end-of-day form, as a full day', () => {
    const ws = readSheet('<row r="1"><c r="A1" t="d"><v>24:00:00</v></c></row>');
    expect(getCell(ws, 1, 1)?.value).toEqual({ kind: 'duration', ms: 86_400_000 });
  });

  it('rejects an hour past the end of the day', () => {
    expect(() => readSheet('<row r="1"><c r="A1" t="d"><v>24:00:01</v></c></row>')).toThrow(
      OpenXmlSchemaError,
    );
    expect(() => readSheet('<row r="1"><c r="A1" t="d"><v>25:00:00</v></c></row>')).toThrow(
      OpenXmlSchemaError,
    );
  });

  it('reads a duration carrying days', () => {
    const ws = readSheet(
      '<row r="1"><c r="A1" t="d"><v>P1D</v></c><c r="B1" t="d"><v>P1DT2H30M</v></c></row>',
    );
    expect(getCell(ws, 1, 1)?.value).toEqual({ kind: 'duration', ms: 86_400_000 });
    expect(getCell(ws, 1, 2)?.value).toEqual({ kind: 'duration', ms: 95_400_000 });
  });

  it('rejects a duration in years or months, which stand for no fixed span', () => {
    expect(() => readSheet('<row r="1"><c r="A1" t="d"><v>P1Y</v></c></row>')).toThrow(OpenXmlSchemaError);
    expect(() => readSheet('<row r="1"><c r="A1" t="d"><v>P2M</v></c></row>')).toThrow(OpenXmlSchemaError);
    expect(() => readSheet('<row r="1"><c r="A1" t="d"><v>P</v></c></row>')).toThrow(OpenXmlSchemaError);
  });

  it('reads a blank value as an empty cell, as t="n" and t="b" do', () => {
    const ws = readSheet('<row r="1"><c r="A1" t="d"><v></v></c><c r="B1" t="d"/></row>');
    expect(getCell(ws, 1, 1)?.value).toBeNull();
    expect(getCell(ws, 1, 2)?.value).toBeNull();
  });

  it('rejects a date that does not exist', () => {
    expect(() => readSheet('<row r="1"><c r="A1" t="d"><v>2024-02-30</v></c></row>')).toThrow(
      'worksheet: <v>2024-02-30</v> at Data!A1 is not an ISO 8601 date, time or duration',
    );
  });

  it('rejects text', () => {
    expect(() => readSheet('<row r="1"><c r="A1" t="d"><v>last tuesday</v></c></row>')).toThrow(
      OpenXmlSchemaError,
    );
  });

  it('loads through loadWorkbook and saves back as a serial, not as t="d"', async () => {
    const { wb, ws } = await loadPatched('<c r="A1" t="d"><v>2024-03-14T00:00:00</v></c>');
    expect(getCell(ws, 1, 1)?.value).toEqual(new Date(Date.UTC(2024, 2, 14)));
    const out = await savedSheetText(wb);
    expect(out).toContain('<c r="A1"><v>45365</v></c>');
    expect(out).not.toContain('t="d"');
  });
});

describe('a cell type outside ST_CellType', () => {
  it('is still refused, naming the type', () => {
    expect(() => readSheet('<row r="1"><c r="A1" t="q"><v>1</v></c></row>')).toThrow(
      'worksheet: unknown cell type t="q"',
    );
  });
});

describe('ISO cell boundary regressions', () => {
  it.each(['2024-03-14T00:00:00+99:00', '2024-03-14T00:00:00+14:01',
    '2024-03-14T00:00:00+02:60', '12:30:00-15:00', '12:30:00+00:60',
    '24:00:00.0001', 'P1DT', `PT${'9'.repeat(310)}S`])('rejects %s before save', (value) => {
    expect(() => readSheet(`<row r="1"><c r="A1" t="d"><v>${value}</v></c></row>`)).toThrow(OpenXmlSchemaError);
  });

  it('accepts a leap day in a year below 100 without applying the 1900 calendar', () => {
    const ws = readSheet('<row r="1"><c r="A1" t="d"><v>0040-02-29</v></c></row>');
    const value = getCell(ws, 1, 1)?.value;
    expect(value instanceof Date ? value.toISOString() : value).toBe('0040-02-29T00:00:00.000Z');
  });

  it.each([false, true])('keeps a formula ISO date numeric on save with date1904=%s', async (date1904) => {
    const bytes = await savedWithCell('<c r="A1" t="d"><f>DATE(2024,3,14)</f><v>2024-03-14T00:00:00</v></c>', date1904);
    const wb = await loadWorkbook(fromBuffer(bytes));
    const serial = date1904 ? 43903 : 45365;
    const ws = getSheet(wb, 'Data');
    if (!ws) throw new Error('missing Data sheet');
    expect(getCell(ws, 1, 1)?.value).toMatchObject({ kind: 'formula', cachedValue: serial });
    const saved = await workbookToBytes(wb);
    const reloaded = await loadWorkbook(fromBuffer(saved));
    const reloadedSheet = getSheet(reloaded, 'Data');
    if (!reloadedSheet) throw new Error('missing reloaded Data sheet');
    expect(getCell(reloadedSheet, 1, 1)?.value).toMatchObject({ cachedValue: serial });
    expect(new TextDecoder().decode(unzipSync(saved)[SHEET_PART])).not.toContain('t="str"');
  });

  it('keeps a formula duration cache numeric on save', async () => {
    const { wb, ws } = await loadPatched('<c r="A1" t="d"><f>A2</f><v>PT12H</v></c>');
    expect(getCell(ws, 1, 1)?.value).toMatchObject({ cachedValue: 0.5 });
    expect(await savedSheetText(wb)).toContain('<f>A2</f><v>0.5</v>');
  });

  it.each(['<v>oops</v>', ''])('validates an unknown formula cell type with %s', (value) => {
    expect(() => readSheet(`<row r="1"><c r="A1" t="q"><f>A2</f>${value}</c></row>`)).toThrow(OpenXmlSchemaError);
  });

  it('rejects a malformed date formula cache during load', () => {
    expect(() => readSheet('<row r="1"><c r="A1" t="d"><f>A2</f><v>oops</v></c></row>')).toThrow(OpenXmlSchemaError);
  });
});
