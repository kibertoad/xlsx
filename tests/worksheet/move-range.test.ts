// Tests for moveRange.

import { describe, expect, it } from 'vitest';
import { addWorksheet, createWorkbook } from '../../src/workbook/workbook.js';
import {
  clearRange,
  copyRange,
  getCell,
  getRangeValues,
  moveRange,
  setCell,
} from '../../src/worksheet/worksheet.js';

describe('moveRange', () => {
  it('moves cells to a non-overlapping target and clears the source', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setCell(ws, 1, 1, 'a');
    setCell(ws, 1, 2, 'b');
    setCell(ws, 2, 1, 'c');
    expect(moveRange(ws, 'A1:B2', 'D5:E6')).toBe(3);
    expect(ws.rows.get(1)).toBeUndefined();
    expect(ws.rows.get(2)).toBeUndefined();
    expect(getRangeValues(ws, 'D5:E6')).toEqual([
      ['a', 'b'],
      ['c', null],
    ]);
  });

  it('overlapping move keeps the destination intact', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setCell(ws, 1, 1, 'a');
    setCell(ws, 1, 2, 'b');
    setCell(ws, 1, 3, 'c');
    // Shift A1:C1 → B1:D1 (overlap = B1, C1).
    moveRange(ws, 'A1:C1', 'B1:D1');
    expect(ws.rows.get(1)?.get(1)).toBeUndefined();
    expect(ws.rows.get(1)?.get(2)?.value).toBe('a');
    expect(ws.rows.get(1)?.get(3)?.value).toBe('b');
    expect(ws.rows.get(1)?.get(4)?.value).toBe('c');
  });

  it('cross-worksheet move via opts.targetWs clears source on the source sheet', () => {
    const wb = createWorkbook();
    const a = addWorksheet(wb, 'A');
    const b = addWorksheet(wb, 'B');
    setCell(a, 1, 1, 'src');
    expect(moveRange(a, 'A1', 'C3', { targetWs: b })).toBe(1);
    expect(a.rows.get(1)).toBeUndefined();
    expect(b.rows.get(3)?.get(3)?.value).toBe('src');
  });

  it('returns 0 for empty source range', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    expect(moveRange(ws, 'A1:B2', 'D5:E6')).toBe(0);
  });

  it('blanks a destination cell whose source cell is empty', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setCell(ws, 1, 1, 1);
    setCell(ws, 3, 1, 3);
    setCell(ws, 2, 3, 'in the way');
    expect(moveRange(ws, 'A1:A3', 'C1:C3')).toBe(2);
    expect(getRangeValues(ws, 'C1:C3')).toEqual([[1], [null], [3]]);
    expect(getRangeValues(ws, 'A1:A3')).toEqual([[null], [null], [null]]);
  });

  it('moves a sparse source into an overlapping target', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setCell(ws, 1, 1, 1);
    setCell(ws, 3, 1, 3);
    expect(moveRange(ws, 'A1:A3', 'A2:A4')).toBe(2);
    expect(getRangeValues(ws, 'A1:A4')).toEqual([[null], [1], [null], [3]]);
  });

  it('is not the same as copyRange followed by clearing the source', () => {
    // Clearing afterwards would take the cells the move had just landed in the
    // overlap, which is why moveRange reads the whole band up front.
    const wb = createWorkbook();
    const moved = addWorksheet(wb, 'moved');
    const copied = addWorksheet(wb, 'copied');
    for (const ws of [moved, copied]) {
      setCell(ws, 1, 1, 1);
      setCell(ws, 2, 1, 2);
      setCell(ws, 3, 1, 3);
    }
    moveRange(moved, 'A1:A3', 'A2:A4');
    copyRange(copied, 'A1:A3', 'A2:A4');
    clearRange(copied, 'A1:A3');
    expect(getRangeValues(moved, 'A1:A4')).toEqual([[null], [1], [2], [3]]);
    expect(getRangeValues(copied, 'A1:A4')).toEqual([[null], [null], [null], [3]]);
  });

  it('does not leave the destination wearing its old hyperlink or comment', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setCell(ws, 1, 1, 'plain');
    const b1 = setCell(ws, 1, 2, 'linked');
    b1.hyperlinkId = 9;
    b1.commentId = 4;
    moveRange(ws, 'A1', 'B1');
    const landed = getCell(ws, 1, 2);
    expect(landed?.value).toBe('plain');
    expect(landed?.hyperlinkId).toBeUndefined();
    expect(landed?.commentId).toBeUndefined();
  });
});
