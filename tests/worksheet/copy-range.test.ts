// Tests for copyRange.

import { describe, expect, it } from 'vitest';
import { setCellFont } from '../../src/styles/cell-style.js';
import { makeFont } from '../../src/styles/fonts.js';
import { addWorksheet, createWorkbook } from '../../src/workbook/workbook.js';
import {
  copyRange,
  getCell,
  getRangeValues,
  setCell,
} from '../../src/worksheet/worksheet.js';

describe('copyRange', () => {
  it('copies cell values to a new top-left within the same sheet', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setCell(ws, 1, 1, 'a');
    setCell(ws, 1, 2, 'b');
    setCell(ws, 2, 1, 'c');
    setCell(ws, 2, 2, 'd');
    expect(copyRange(ws, 'A1:B2', 'D5:E6')).toBe(4);
    expect(getRangeValues(ws, 'D5:E6')).toEqual([
      ['a', 'b'],
      ['c', 'd'],
    ]);
    // Source unchanged.
    expect(getRangeValues(ws, 'A1:B2')).toEqual([
      ['a', 'b'],
      ['c', 'd'],
    ]);
  });

  it('preserves styleId on copied cells (shared stylesheet)', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    const orig = setCell(ws, 1, 1, 'x');
    setCellFont(wb, orig, makeFont({ bold: true }));
    copyRange(ws, 'A1', 'C3');
    const copy = ws.rows.get(3)?.get(3);
    if (!copy) throw new Error('expected C3');
    expect(copy.styleId).toBe(orig.styleId);
    expect(copy.value).toBe('x');
  });

  it('counts only the populated source cells', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setCell(ws, 1, 1, 'a');
    setCell(ws, 2, 2, 'd');
    // A1:B2 has only 2 populated; A1, B2.
    expect(copyRange(ws, 'A1:B2', 'D1:E2')).toBe(2);
    expect(ws.rows.get(1)?.get(4)?.value).toBe('a');
    expect(ws.rows.get(2)?.get(5)?.value).toBe('d');
  });

  it('truncates when target range is smaller than source', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setCell(ws, 1, 1, 'a');
    setCell(ws, 1, 2, 'b');
    setCell(ws, 1, 3, 'c');
    expect(copyRange(ws, 'A1:C1', 'E1:F1')).toBe(2);
    expect(ws.rows.get(1)?.get(5)?.value).toBe('a');
    expect(ws.rows.get(1)?.get(6)?.value).toBe('b');
    expect(ws.rows.get(1)?.get(7)).toBeUndefined();
  });

  it('cross-worksheet copy via opts.targetWs', () => {
    const wb = createWorkbook();
    const a = addWorksheet(wb, 'A');
    const b = addWorksheet(wb, 'B');
    setCell(a, 1, 1, 'src');
    expect(copyRange(a, 'A1', 'C3', { targetWs: b })).toBe(1);
    expect(b.rows.get(3)?.get(3)?.value).toBe('src');
    expect(a.rows.get(3)?.get(3)).toBeUndefined();
  });

  it('copies down into an overlapping target without re-reading its own writes', () => {
    // A1:A3 shifted one row down. The write loop used to read `ws.rows` as it
    // went, so A2 was already the copy of A1 by the time it was read, and
    // [1, 2, 3] collapsed to [1, 1, 1].
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setCell(ws, 1, 1, 1);
    setCell(ws, 2, 1, 2);
    setCell(ws, 3, 1, 3);
    expect(copyRange(ws, 'A1:A3', 'A2:A4')).toBe(3);
    expect(getRangeValues(ws, 'A1:A4')).toEqual([[1], [1], [2], [3]]);
  });

  it('copies up into an overlapping target', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setCell(ws, 2, 1, 1);
    setCell(ws, 3, 1, 2);
    setCell(ws, 4, 1, 3);
    expect(copyRange(ws, 'A2:A4', 'A1:A3')).toBe(3);
    expect(getRangeValues(ws, 'A1:A4')).toEqual([[1], [2], [3], [3]]);
  });

  it('copies right into an overlapping target', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setCell(ws, 1, 1, 1);
    setCell(ws, 1, 2, 2);
    setCell(ws, 1, 3, 3);
    expect(copyRange(ws, 'A1:C1', 'B1:D1')).toBe(3);
    expect(getRangeValues(ws, 'A1:D1')).toEqual([[1, 1, 2, 3]]);
  });

  it('copies onto itself as a no-op', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setCell(ws, 1, 1, 1);
    setCell(ws, 2, 1, 2);
    expect(copyRange(ws, 'A1:A2', 'A1:A2')).toBe(2);
    expect(getRangeValues(ws, 'A1:A2')).toEqual([[1], [2]]);
  });

  it('carries a hyperlinkId through an overlapping same-sheet copy', () => {
    // The id is read off a detached copy taken before the first write, so it
    // still travels when the source coordinate is one the copy lands on.
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    const a1 = setCell(ws, 1, 1, 'link');
    a1.hyperlinkId = 7;
    setCell(ws, 2, 1, 'plain');
    copyRange(ws, 'A1:A2', 'A2:A3');
    expect(ws.rows.get(2)?.get(1)?.hyperlinkId).toBe(7);
    expect(ws.rows.get(3)?.get(1)?.hyperlinkId).toBeUndefined();
  });

  it('overwrites existing cells in the target extent', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setCell(ws, 1, 1, 'src');
    setCell(ws, 5, 5, 'dst-original');
    copyRange(ws, 'A1', 'E5');
    expect(ws.rows.get(5)?.get(5)?.value).toBe('src');
  });

  it('blanks a destination cell whose source cell is empty', () => {
    // Pasting a range in Excel replaces the whole landing rectangle. A gap in
    // the source has to land as a gap, not leave the destination's own value
    // sitting in the middle of the pasted block.
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setCell(ws, 1, 1, 1);
    setCell(ws, 3, 1, 3);
    setCell(ws, 2, 3, 'in the way');
    expect(copyRange(ws, 'A1:A3', 'C1:C3')).toBe(2);
    expect(getRangeValues(ws, 'C1:C3')).toEqual([[1], [null], [3]]);
  });

  it('blanks the gap when a sparse source copies into an overlapping target', () => {
    // A1 and A3 populated, A2 empty, shifted one row down. The gap has to
    // reach A3 and evict the 3 that lives there; returning 2 rather than 3
    // follows from A2 never being read as populated.
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setCell(ws, 1, 1, 1);
    setCell(ws, 3, 1, 3);
    expect(copyRange(ws, 'A1:A3', 'A2:A4')).toBe(2);
    expect(getRangeValues(ws, 'A1:A4')).toEqual([[1], [1], [null], [3]]);
  });

  it('does not leave the destination wearing its old hyperlink or comment', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setCell(ws, 1, 1, 'plain');
    const b1 = setCell(ws, 1, 2, 'linked');
    b1.hyperlinkId = 9;
    b1.commentId = 4;
    copyRange(ws, 'A1', 'B1');
    const landed = getCell(ws, 1, 2);
    expect(landed?.value).toBe('plain');
    expect(landed?.hyperlinkId).toBeUndefined();
    expect(landed?.commentId).toBeUndefined();
  });

  it('clears a stale id inside an overlapping copy', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setCell(ws, 1, 1, 1);
    const a2 = setCell(ws, 2, 1, 2);
    a2.commentId = 42;
    setCell(ws, 3, 1, 3);
    copyRange(ws, 'A1:A3', 'A2:A4');
    expect(getCell(ws, 2, 1)?.value).toBe(1);
    expect(getCell(ws, 2, 1)?.commentId).toBeUndefined();
  });

  it('reads the source sheet and writes the target sheet, never the reverse', () => {
    // The detach reads `ws` while the replay writes `dest`; mixing the two up
    // only shows on a cross-sheet copy.
    const wb = createWorkbook();
    const a = addWorksheet(wb, 'A');
    const b = addWorksheet(wb, 'B');
    setCell(a, 1, 1, 'a1');
    setCell(a, 2, 2, 'b2');
    setCell(b, 1, 1, 'clobber me');
    expect(copyRange(a, 'A1:B2', 'A1:B2', { targetWs: b })).toBe(2);
    expect(getRangeValues(b, 'A1:B2')).toEqual([
      ['a1', null],
      [null, 'b2'],
    ]);
    // Source sheet untouched.
    expect(getRangeValues(a, 'A1:B2')).toEqual([
      ['a1', null],
      [null, 'b2'],
    ]);
  });

  it('drops hyperlinkId across sheets and inherits nothing from the destination', () => {
    const wb = createWorkbook();
    const a = addWorksheet(wb, 'A');
    const b = addWorksheet(wb, 'B');
    setCell(a, 1, 1, 'src');
    const stale = setCell(b, 1, 1, 'dst');
    stale.hyperlinkId = 3;
    stale.commentId = 4;
    copyRange(a, 'A1', 'A1', { targetWs: b });
    const landed = getCell(b, 1, 1);
    expect(landed?.value).toBe('src');
    expect(landed?.hyperlinkId).toBeUndefined();
    expect(landed?.commentId).toBeUndefined();
  });
});
