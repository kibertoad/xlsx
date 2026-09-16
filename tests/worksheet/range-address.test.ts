// Tests for getRangeAddress — sheet-qualified A1 range string.

import { describe, expect, it } from 'vitest';
import { OpenXmlSchemaError } from '../../src/utils/exceptions.js';
import { addWorksheet, createWorkbook } from '../../src/workbook/workbook.js';
import { getRangeAddress } from '../../src/worksheet/worksheet.js';

describe('getRangeAddress', () => {
  it('returns the sheet-qualified address for a single-cell ref', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'Data');
    expect(getRangeAddress(ws, 'A1')).toBe('Data!A1');
  });

  it('returns the sheet-qualified address for a rectangular range', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'Data');
    expect(getRangeAddress(ws, 'A1:B5')).toBe('Data!A1:B5');
  });

  it('quotes the title when it contains spaces / punctuation', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'Q1 2024');
    expect(getRangeAddress(ws, 'A1:B5')).toBe("'Q1 2024'!A1:B5");
  });

  it('keeps a row span, a column span and $ markers as written', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'S');
    expect(getRangeAddress(ws, '1:5')).toBe('S!1:5');
    expect(getRangeAddress(ws, 'A:E')).toBe('S!A:E');
    // The absolute form is what a defined name wants; normalising would drop
    // the markers and expand the span.
    expect(getRangeAddress(ws, '$A$4:$H$20')).toBe('S!$A$4:$H$20');
  });

  it('rejects a string that is not a range, like it rejects bad bounds', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'S');
    // Forwarded verbatim, "Data!not a range" reaches workbook.xml through
    // addDefinedNameForRange and Excel asks to repair the file on open.
    expect(() => getRangeAddress(ws, 'not a range')).toThrow(OpenXmlSchemaError);
    expect(() => getRangeAddress(ws, 'A0')).toThrow(OpenXmlSchemaError);
    expect(() => getRangeAddress(ws, '')).toThrow(OpenXmlSchemaError);
  });
});
