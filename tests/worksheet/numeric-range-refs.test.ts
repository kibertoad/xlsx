// Tests for passing numeric bounds where an A1 range string is accepted.
//
// A caller that already tracks rows and columns as integers should not have to
// format a string for the callee to parse straight back.

import { describe, expect, it } from 'vitest';
import { makeBorder, makeSide } from '../../src/styles/borders.js';
import {
  getCellBorder,
  getCellNumberFormat,
  setRangeNumberFormat,
  setRangeStyle,
} from '../../src/styles/cell-style.js';
import { MAX_COL, MAX_ROW } from '../../src/utils/coordinate.js';
import { OpenXmlSchemaError } from '../../src/utils/exceptions.js';
import { addWorksheet, createWorkbook } from '../../src/workbook/workbook.js';
import {
  applyToRange,
  clearRange,
  copyRange,
  getCell,
  getMergedCells,
  getRangeAddress,
  getRangeValues,
  mergeCells,
  moveRange,
  setCell,
  setRangeValues,
  writeRange,
} from '../../src/worksheet/worksheet.js';

const bounds = { minRow: 2, minCol: 2, maxRow: 3, maxCol: 3 };

describe('numeric bounds as a range ref', () => {
  it('setRangeValues / getRangeValues agree with the A1 form', () => {
    const wb = createWorkbook();
    const numeric = addWorksheet(wb, 'N');
    const a1 = addWorksheet(wb, 'S');
    const values = [
      [1, 2],
      [3, 4],
    ];
    setRangeValues(numeric, bounds, values);
    setRangeValues(a1, 'B2:C3', values);
    expect(getRangeValues(numeric, bounds)).toEqual(getRangeValues(a1, 'B2:C3'));
    expect(getRangeValues(numeric, bounds)).toEqual(values);
  });

  it('applyToRange visits every coordinate and allocates on first touch', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    const seen: string[] = [];
    applyToRange(ws, bounds, (_cell, row, col) => seen.push(`${row},${col}`));
    expect(seen).toEqual(['2,2', '2,3', '3,2', '3,3']);
    expect(getCell(ws, 3, 3)?.value).toBeNull();
  });

  it('clearRange counts the same cells either way', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setRangeValues(ws, bounds, [
      [1, 2],
      [3, 4],
    ]);
    expect(clearRange(ws, bounds)).toBe(4);
  });

  it('style helpers accept bounds', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setRangeNumberFormat(wb, ws, bounds, '#,##0');
    setRangeStyle(wb, ws, bounds, { border: makeBorder({ top: makeSide({ style: 'thin' }) }) });
    const c = getCell(ws, 2, 2);
    expect(c).toBeDefined();
    if (!c) return;
    expect(getCellNumberFormat(wb, c)).toBe('#,##0');
    expect(getCellBorder(wb, c).top?.style).toBe('thin');
  });

  it('normalises inverted bounds like the A1 parser does', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setCell(ws, 1, 1, 'x');
    setCell(ws, 2, 2, 'y');
    expect(clearRange(ws, { minRow: 2, minCol: 2, maxRow: 1, maxCol: 1 })).toBe(2);
  });

  it('writeRange takes a numeric anchor', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    const written = writeRange(ws, { row: 2, col: 3 }, [['a', 'b']]);
    expect(written).toEqual({ minRow: 2, maxRow: 2, minCol: 3, maxCol: 4 });
    expect(getCell(ws, 2, 4)?.value).toBe('b');
  });

  it('getRangeAddress formats bounds as a sheet-qualified rectangle', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'Quarter 1');
    expect(getRangeAddress(ws, bounds)).toBe("'Quarter 1'!B2:C3");
  });

  it('copyRange / moveRange take bounds on both sides', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    const target = { minRow: 6, minCol: 2, maxRow: 7, maxCol: 3 };
    setRangeValues(ws, bounds, [
      [1, 2],
      [3, 4],
    ]);
    expect(copyRange(ws, bounds, target)).toBe(4);
    expect(getRangeValues(ws, target)).toEqual([
      [1, 2],
      [3, 4],
    ]);
    expect(moveRange(ws, bounds, { minRow: 10, minCol: 2, maxRow: 11, maxCol: 3 })).toBe(4);
    expect(getRangeValues(ws, bounds)).toEqual([
      [null, null],
      [null, null],
    ]);
  });
});

// The A1 parser rejects "A0", a fractional row, and anything past the grid
// ceiling. Numeric bounds reach the same helpers, so they have to fail the same
// way rather than iterate a region Excel cannot express.
describe('numeric bounds are validated, not trusted', () => {
  const invalid = [
    { minRow: 0, minCol: 0, maxRow: 0, maxCol: 0 },
    { minRow: 1.5, minCol: 1, maxRow: 3.5, maxCol: 3 },
    { minRow: Number.NaN, minCol: 1, maxRow: 3, maxCol: 3 },
    { minRow: 1, minCol: 1, maxRow: MAX_ROW + 1, maxCol: 3 },
    { minRow: 1, minCol: 1, maxRow: 3, maxCol: MAX_COL + 1 },
    { minRow: 1, minCol: 1, maxRow: 1e15, maxCol: 3 },
  ];

  it('every range consumer throws instead of silently doing nothing', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    for (const bad of invalid) {
      expect(() => clearRange(ws, bad)).toThrow(OpenXmlSchemaError);
      expect(() => getRangeValues(ws, bad)).toThrow(OpenXmlSchemaError);
      expect(() => applyToRange(ws, bad, () => {})).toThrow(OpenXmlSchemaError);
      expect(() => setRangeNumberFormat(wb, ws, bad, '#,##0')).toThrow(OpenXmlSchemaError);
      expect(() => getRangeAddress(ws, bad)).toThrow(OpenXmlSchemaError);
    }
    expect(ws.rows.size).toBe(0);
  });

  it('writeRange checks a numeric anchor even when every value is skipped', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    expect(() => writeRange(ws, { row: 0, col: -3 }, [[null]])).toThrow(OpenXmlSchemaError);
    expect(() => writeRange(ws, { row: 1, col: 1.5 }, [['a']])).toThrow(OpenXmlSchemaError);
    expect(ws.rows.size).toBe(0);
  });
});

describe('bounds define the region, not just its corner', () => {
  it('setRangeValues drops values past the range instead of writing outside it', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    const oneCell = { minRow: 1, minCol: 1, maxRow: 1, maxCol: 1 };
    setRangeValues(ws, oneCell, [
      [1, 2, 3],
      [4, 5, 6],
    ]);
    expect(getRangeValues(ws, oneCell)).toEqual([[1]]);
    expect(getCell(ws, 1, 2)).toBeUndefined();
    expect(getCell(ws, 2, 1)).toBeUndefined();
  });

  it('mergeCells normalises the bounds and keeps its own copy', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    const inverted = { minRow: 5, minCol: 3, maxRow: 1, maxCol: 1 };
    expect(mergeCells(ws, inverted)).toEqual({ minRow: 1, minCol: 1, maxRow: 5, maxCol: 3 });
    inverted.maxRow = 500;
    expect(getMergedCells(ws)[0]).toEqual({ minRow: 1, minCol: 1, maxRow: 5, maxCol: 3 });
  });

  it('getRangeAddress names the region the other helpers operate on', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'Quarter 1');
    const inverted = { minRow: 5, minCol: 1, maxRow: 1, maxCol: 1 };
    setCell(ws, 3, 1, 'inside A1:A5');
    expect(getRangeAddress(ws, inverted)).toBe("'Quarter 1'!A1:A5");
    expect(clearRange(ws, inverted)).toBe(1);
  });
});
