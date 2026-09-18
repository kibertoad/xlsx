// getSheet / getChartsheet / getSheetByIndex all discriminate the sheet union,
// so a title or index naming the other kind resolves to undefined rather than
// to a sheet the caller cannot use.

import { describe, expect, it } from 'vitest';
import {
  addChartsheet,
  addWorksheet,
  createWorkbook,
  getChartsheet,
  getSheet,
  getSheetByIndex,
} from '../../src/workbook/workbook.js';

describe('sheet lookup discriminates worksheets from chartsheets', () => {
  it('getSheet resolves a worksheet title and misses a chartsheet one', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'Data');
    addChartsheet(wb, 'Chart1');
    expect(getSheet(wb, 'Data')).toBe(ws);
    expect(getSheet(wb, 'Chart1')).toBeUndefined();
    expect(getSheet(wb, 'Missing')).toBeUndefined();
  });

  it('getChartsheet resolves a chartsheet title and misses a worksheet one', () => {
    const wb = createWorkbook();
    addWorksheet(wb, 'Data');
    const cs = addChartsheet(wb, 'Chart1');
    expect(getChartsheet(wb, 'Chart1')).toBe(cs);
    expect(getChartsheet(wb, 'Data')).toBeUndefined();
    expect(getChartsheet(wb, 'Missing')).toBeUndefined();
  });

  it('getSheetByIndex returns undefined for a chartsheet slot', () => {
    const wb = createWorkbook();
    addChartsheet(wb, 'Chart1');
    const ws = addWorksheet(wb, 'Data');
    expect(getSheetByIndex(wb, 0)).toBeUndefined();
    expect(getSheetByIndex(wb, 1)).toBe(ws);
    expect(getSheetByIndex(wb, 2)).toBeUndefined();
    expect(getSheetByIndex(wb, -1)).toBeUndefined();
  });

  it('an empty workbook resolves nothing', () => {
    const wb = createWorkbook();
    expect(getSheet(wb, 'anything')).toBeUndefined();
    expect(getChartsheet(wb, 'anything')).toBeUndefined();
    expect(getSheetByIndex(wb, 0)).toBeUndefined();
  });
});
