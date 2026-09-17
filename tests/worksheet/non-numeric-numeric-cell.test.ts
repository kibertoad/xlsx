// A numeric cell has to carry a finite number: that is what the cell model can
// hold and what the writer can emit. These cover both ends, so the reader and
// the writer cannot drift apart on which files are loadable.

import { unzipSync, zipSync } from 'fflate';
import { describe, expect, it } from 'vitest';
import { setFormula } from '../../src/cell/cell.js';
import { loadWorkbook } from '../../src/io/load.js';
import { fromBuffer } from '../../src/io/node.js';
import { workbookToBytes } from '../../src/io/save.js';
import { OpenXmlSchemaError } from '../../src/utils/exceptions.js';
import { addWorksheet, createWorkbook } from '../../src/workbook/workbook.js';
import { parseWorksheetXml } from '../../src/worksheet/reader.js';
import { ensureCell, getCell, setCell } from '../../src/worksheet/worksheet.js';

const MAIN_NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const SHEET_PART = 'xl/worksheets/sheet1.xml';

const sheet = (body: string): string =>
  `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
  `<worksheet xmlns="${MAIN_NS}"><sheetData>${body}</sheetData></worksheet>`;

const readSheet = (body: string) => parseWorksheetXml(sheet(body), 'Data', { sharedStrings: [] });

/**
 * A workbook saved by the library, with the `<v>` text of its one numeric cell
 * swapped out. Going through the real save path keeps the rest of the package
 * (content types, rels, workbook part) exactly as a legitimate file has it, so
 * the load fails on the cell and nothing else.
 */
const savedWithCellText = async (text: string): Promise<Uint8Array> => {
  const wb = createWorkbook();
  setCell(addWorksheet(wb, 'Data'), 1, 1, 42);
  const entries = unzipSync(await workbookToBytes(wb));
  const part = entries[SHEET_PART];
  if (part === undefined) throw new Error(`save produced no ${SHEET_PART}`);
  const patched = new TextDecoder().decode(part).replace('<v>42</v>', `<v>${text}</v>`);
  entries[SHEET_PART] = new TextEncoder().encode(patched);
  return zipSync(entries);
};

describe('a numeric cell whose value is not a finite number', () => {
  it('names the sheet, the cell and the text it could not read', () => {
    expect(() => readSheet('<row r="2"><c r="B2"><v>oops</v></c></row>')).toThrow(OpenXmlSchemaError);
    expect(() => readSheet('<row r="2"><c r="B2"><v>oops</v></c></row>')).toThrow(
      'worksheet: <v>oops</v> at Data!B2 is not a finite number',
    );
  });

  it('rejects an exponent past the double range, through loadWorkbook', async () => {
    expect(() => readSheet('<row r="1"><c r="A1"><v>1e400</v></c></row>')).toThrow(
      'worksheet: <v>1e400</v> at Data!A1 is not a finite number',
    );
    await expect(loadWorkbook(fromBuffer(await savedWithCellText('1e400')))).rejects.toThrow(
      'worksheet: <v>1e400</v> at Data!A1 is not a finite number',
    );
  });

  it('rejects text, through loadWorkbook', async () => {
    await expect(loadWorkbook(fromBuffer(await savedWithCellText('oops')))).rejects.toThrow(
      OpenXmlSchemaError,
    );
  });

  it('rejects it in a cached formula result too', () => {
    expect(() => readSheet('<row r="1"><c r="C1"><f>A2</f><v>oops</v></c></row>')).toThrow(
      'worksheet: <v>oops</v> at Data!C1 is not a finite number',
    );
  });

  it('quotes a bounded prefix of the offending text', () => {
    const long = 'x'.repeat(5000);
    let message = '';
    try {
      readSheet(`<row r="1"><c r="A1"><v>${long}</v></c></row>`);
    } catch (err) {
      message = (err as Error).message;
    }
    expect(message).toContain('(5000 chars)');
    expect(message.length).toBeLessThan(120);
  });
});

describe('a numeric cell with no number in it', () => {
  it('reads as an empty cell, whitespace included', () => {
    // The XML reader keeps element text verbatim, so a pretty-printed part
    // reaches the cell readers with the indentation still attached.
    const ws = readSheet(
      '<row r="1">' +
        '<c r="A1"><v></v></c>' +
        '<c r="B1"><v> </v></c>' +
        '<c r="C1"><v>\n      </v></c>' +
        '<c r="D1"/>' +
        '</row>',
    );
    for (const col of [1, 2, 3, 4]) expect(getCell(ws, 1, col)?.value).toBeNull();
  });

  it('reads a blank cached formula result as no cached value', () => {
    const ws = readSheet('<row r="1"><c r="A1"><f>A2</f><v>\n</v></c></row>');
    expect(getCell(ws, 1, 1)?.value).toEqual({ kind: 'formula', t: 'normal', formula: 'A2' });
  });

  it('leaves ordinary numbers alone', () => {
    const ws = readSheet(
      '<row r="1"><c r="A1"><v>42</v></c><c r="B1"><v>-3.5e-7</v></c><c r="C1"><v> 7 </v></c></row>',
    );
    expect(getCell(ws, 1, 1)?.value).toBe(42);
    expect(getCell(ws, 1, 2)?.value).toBeCloseTo(-3.5e-7);
    expect(getCell(ws, 1, 3)?.value).toBe(7);
  });
});

describe('the writer refuses what the reader refuses', () => {
  it('rejects a non-finite cached formula value instead of writing <v>NaN</v>', async () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'Data');
    setFormula(ensureCell(ws, 1, 1), 'A2', { cachedValue: Number.NaN });
    await expect(workbookToBytes(wb)).rejects.toThrow('cannot serialise non-finite number at A1');
  });

  it('rejects a non-finite cell value', async () => {
    const wb = createWorkbook();
    setCell(addWorksheet(wb, 'Data'), 3, 1, Number.POSITIVE_INFINITY);
    await expect(workbookToBytes(wb)).rejects.toThrow('cannot serialise non-finite number at A3');
  });
});
