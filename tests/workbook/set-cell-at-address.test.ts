// Tests for setCellAtAddress — sheet-qualified A1 → single-cell write.

import { describe, expect, it } from 'vitest';
import { OpenXmlSchemaError } from '../../src/utils/exceptions.js';
import {
  addWorksheet,
  createWorkbook,
  getCellAtAddress,
  setCellAtAddress,
} from '../../src/workbook/workbook.js';

describe('setCellAtAddress', () => {
  it('writes a value at the address and returns the Cell', () => {
    const wb = createWorkbook();
    addWorksheet(wb, 'Data');
    const cell = setCellAtAddress(wb, 'Data!B2', 'hello');
    expect(cell.row).toBe(2);
    expect(cell.col).toBe(2);
    expect(cell.value).toBe('hello');
    expect(getCellAtAddress(wb, 'Data!B2')?.value).toBe('hello');
  });

  it('honours quoted sheet titles', () => {
    const wb = createWorkbook();
    addWorksheet(wb, 'Q1 2024');
    setCellAtAddress(wb, "'Q1 2024'!A1", 42);
    expect(getCellAtAddress(wb, "'Q1 2024'!A1")?.value).toBe(42);
  });

  it('throws when the sheet does not exist', () => {
    const wb = createWorkbook();
    expect(() => setCellAtAddress(wb, 'Missing!A1', 'x')).toThrow(/sheet/);
  });

  it('throws when the address points at a range', () => {
    const wb = createWorkbook();
    addWorksheet(wb, 'Data');
    expect(() => setCellAtAddress(wb, 'Data!A1:B5', 'x')).toThrow(/range/);
  });

  // The signature is (wb, address, value), and the mistake it invites is
  // (wb, worksheet, address, value). Before, the Worksheet was coerced into
  // the message as "[object Object]" and blamed for a missing "!".
  it('names the argument when a Worksheet is passed where the address goes', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'Data');
    const call = () => (setCellAtAddress as unknown as (...a: unknown[]) => unknown)(wb, ws, 'A1', 'x');
    expect(call).toThrow(OpenXmlSchemaError);
    expect(call).toThrow(/address must be a sheet-qualified A1 string/);
    expect(call).toThrow(/received an object/);
    expect(call).toThrow(/setCellByCoord/);
    expect(call).not.toThrow(/\[object Object\]/);
  });

  // The mirror mistake: the Worksheet lands in the workbook slot, which used
  // to reach getSheet and die on `wb.sheets` with a bare TypeError.
  it('names the argument when a Worksheet is passed where the workbook goes', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'Data');
    const call = () => (setCellAtAddress as unknown as (...a: unknown[]) => unknown)(ws, 'Data!A1', 'x');
    expect(call).toThrow(OpenXmlSchemaError);
    expect(call).toThrow(/setCellAtAddress: first argument must be the Workbook/);
    expect(call).toThrow(/setCellByCoord/);
    expect(call).not.toThrow(/is not iterable/);
  });

  it('reports a missing address as undefined rather than as a bad string', () => {
    const wb = createWorkbook();
    addWorksheet(wb, 'Data');
    const call = () => (setCellAtAddress as unknown as (...a: unknown[]) => unknown)(wb, undefined, 'x');
    expect(call).toThrow(/received undefined/);
    // No argument was swapped, so the getCellByCoord advice would misdirect.
    expect(call).not.toThrow(/setCellByCoord/);
  });
});
