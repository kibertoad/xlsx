// `getCellDate` composes the three steps reading a date-formatted cell
// otherwise takes: resolve the cell's number format, decide whether it is a
// date format, convert the serial under the workbook epoch.

import { describe, expect, it } from 'vitest';
import type { CellValue } from '../../src/cell/cell.js';
import { makeDurationValue, makeFormula } from '../../src/cell/cell.js';
import { getCellDate, setCellAsDate, setCellNumberFormat } from '../../src/styles/index.js';
import { addWorksheet, createWorkbook } from '../../src/workbook/index.js';
import { setCell } from '../../src/worksheet/index.js';

const read = (code: string, value: CellValue, opts?: { date1904?: boolean }): Date | undefined => {
  const wb = createWorkbook({ date1904: opts?.date1904 ?? false });
  const ws = addWorksheet(wb, 'Sheet1');
  const cell = setCell(ws, 1, 1, value);
  setCellNumberFormat(wb, cell, code);
  return getCellDate(wb, cell);
};

/** 2024-03-14 under the Windows epoch. */
const MARCH_14_2024 = 45_365;

describe('getCellDate', () => {
  it.each(['yyyy-mm-dd', 'mm-dd-yy', 'd-mmm-yy', 'm/d/yy h:mm', 'yyyy-mm-dd hh:mm:ss'])(
    'reads a serial under the date format %s',
    (code) => {
      expect(read(code, MARCH_14_2024)?.toISOString()).toBe('2024-03-14T00:00:00.000Z');
    },
  );

  it('carries the time of day', () => {
    expect(read('yyyy-mm-dd hh:mm:ss', 45_365.7375)?.toISOString()).toBe('2024-03-14T17:42:00.000Z');
  });

  it('uses the workbook epoch', () => {
    expect(read('yyyy-mm-dd', MARCH_14_2024, { date1904: true })?.toISOString()).toBe('2028-03-15T00:00:00.000Z');
  });

  it('reads a time-only format as a time on the epoch day', () => {
    expect(read('h:mm', 0.5)?.toISOString()).toBe('1899-12-31T12:00:00.000Z');
  });

  it('reads the cached value of a formula cell', () => {
    expect(read('yyyy-mm-dd', makeFormula('TODAY()', { cachedValue: MARCH_14_2024 }))?.toISOString()).toBe(
      '2024-03-14T00:00:00.000Z',
    );
  });

  it('passes a Date value through whether or not the cell is formatted', () => {
    const date = new Date(Date.UTC(2024, 2, 14, 9, 30));
    expect(read('yyyy-mm-dd', date)).toBe(date);
    expect(read('General', date)).toBe(date);
  });

  it('pairs with setCellAsDate on the write side', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'Sheet1');
    const cell = setCell(ws, 1, 1, MARCH_14_2024);
    setCellAsDate(wb, cell);
    expect(getCellDate(wb, cell)?.toISOString()).toBe('2024-03-14T00:00:00.000Z');
  });
});

describe('getCellDate: no date reading', () => {
  it('a number with no date format is a count, not a date', () => {
    expect(read('General', MARCH_14_2024)).toBeUndefined();
    expect(read('#,##0.00', MARCH_14_2024)).toBeUndefined();
    expect(read('0%', 0.5)).toBeUndefined();
  });

  it('an elapsed-time format measures a span', () => {
    expect(read('[h]:mm:ss', 1.5)).toBeUndefined();
    expect(read('[mm]:ss', 0.5)).toBeUndefined();
  });

  it('a duration value carries its own span', () => {
    expect(read('[h]:mm', makeDurationValue(5_400_000))).toBeUndefined();
  });

  it.each([
    ['a string', 'text'],
    ['an empty cell', null],
    ['a boolean', true],
  ] satisfies Array<[string, CellValue]>)('%s has no date reading', (_label, value) => {
    expect(read('yyyy-mm-dd', value)).toBeUndefined();
  });

  it('an uncached formula cell has no date reading', () => {
    expect(read('yyyy-mm-dd', makeFormula('TODAY()'))).toBeUndefined();
  });
});
