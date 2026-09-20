// Tests for getCellAtAddress — sheet-qualified A1 → Cell lookup.

import { describe, expect, it } from 'vitest';
import { OpenXmlSchemaError } from '../../src/utils/exceptions.js';
import { addWorksheet, createWorkbook, getCellAtAddress } from '../../src/workbook/workbook.js';
import { getCellAddress, setCell } from '../../src/worksheet/worksheet.js';

describe('getCellAtAddress', () => {
  it('resolves a bare-title address', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'Data');
    setCell(ws, 1, 1, 'a');
    expect(getCellAtAddress(wb, 'Data!A1')?.value).toBe('a');
  });

  it('resolves a quoted-title address', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'Q1 2024');
    setCell(ws, 5, 7, 'cell');
    expect(getCellAtAddress(wb, "'Q1 2024'!G5")?.value).toBe('cell');
  });

  it('returns undefined when the cell is not materialised', () => {
    const wb = createWorkbook();
    addWorksheet(wb, 'Data');
    expect(getCellAtAddress(wb, 'Data!Z99')).toBeUndefined();
  });

  it('throws when the sheet does not exist', () => {
    const wb = createWorkbook();
    expect(() => getCellAtAddress(wb, 'Missing!A1')).toThrow(/sheet/);
  });

  it('throws when the address points at a range instead of a single cell', () => {
    const wb = createWorkbook();
    addWorksheet(wb, 'Data');
    expect(() => getCellAtAddress(wb, 'Data!A1:B5')).toThrow(/range/);
  });

  // The signature is (wb, address), and the mistake it invites is (wb, ws,
  // ref). Before, the Worksheet was coerced into the message as
  // "[object Object]" and blamed for a missing "!".
  it('names the argument when a Worksheet is passed where the address goes', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'Data');
    const call = () => (getCellAtAddress as unknown as (...a: unknown[]) => unknown)(wb, ws, 'A1');
    expect(call).toThrow(OpenXmlSchemaError);
    expect(call).toThrow(/getCellAtAddress: address must be a sheet-qualified A1 string/);
    expect(call).toThrow(/received an object/);
    expect(call).toThrow(/getCellByCoord/);
    expect(call).not.toThrow(/\[object Object\]/);
  });

  // The mirror mistake: the Worksheet lands in the workbook slot, which used
  // to reach getSheet and die on `wb.sheets` with a bare TypeError.
  it('names the argument when a Worksheet is passed where the workbook goes', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'Data');
    const call = () => (getCellAtAddress as unknown as (...a: unknown[]) => unknown)(ws, 'Data!A1');
    expect(call).toThrow(OpenXmlSchemaError);
    expect(call).toThrow(/getCellAtAddress: first argument must be the Workbook/);
    expect(call).toThrow(/getCellByCoord/);
    expect(call).not.toThrow(/is not iterable/);
  });

  it('round-trips through getCellAddress', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'Q1 2024');
    setCell(ws, 3, 4, 'rt');
    const cell = ws.rows.get(3)?.get(4);
    if (!cell) throw new Error('cell missing');
    const address = getCellAddress(ws, cell);
    expect(getCellAtAddress(wb, address)?.value).toBe('rt');
  });
});
