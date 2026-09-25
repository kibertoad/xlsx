// Tests for the defined-name builder helpers.

import { describe, expect, it } from 'vitest';
import { fromBuffer } from '../../src/io/node.js';
import { loadWorkbook } from '../../src/io/load.js';
import { workbookToBytes } from '../../src/io/save.js';
import { addChartsheet, addWorksheet, createWorkbook } from '../../src/workbook/workbook.js';
import {
  addDefinedName,
  getDefinedName,
  getDefinedNameTarget,
  removeDefinedName,
  setPrintArea,
  setPrintTitles,
} from '../../src/workbook/defined-names.js';
import { OpenXmlSchemaError } from '../../src/utils/exceptions.js';

describe('addDefinedName / getDefinedName / removeDefinedName', () => {
  it('addDefinedName replaces an existing entry with the same name + scope', () => {
    const wb = createWorkbook();
    addWorksheet(wb, 'A');
    addDefinedName(wb, { name: 'tax', value: '0.05' });
    expect(wb.definedNames).toHaveLength(1);
    addDefinedName(wb, { name: 'tax', value: '0.08' });
    expect(wb.definedNames).toHaveLength(1);
    expect(wb.definedNames[0]?.value).toBe('0.08');
  });

  it('different scopes coexist (workbook + sheet 0)', () => {
    const wb = createWorkbook();
    addWorksheet(wb, 'A');
    addDefinedName(wb, { name: 'region', value: '"All"' });
    addDefinedName(wb, { name: 'region', value: '"North"', scope: 0 });
    expect(wb.definedNames).toHaveLength(2);
    expect(getDefinedName(wb, 'region')?.value).toBe('"All"');
    expect(getDefinedName(wb, 'region', 0)?.value).toBe('"North"');
  });

  it('removeDefinedName returns true when removed, false otherwise', () => {
    const wb = createWorkbook();
    addDefinedName(wb, { name: 'x', value: '1' });
    expect(removeDefinedName(wb, 'x')).toBe(true);
    expect(removeDefinedName(wb, 'x')).toBe(false);
    expect(wb.definedNames).toHaveLength(0);
  });
});

describe('setPrintArea / setPrintTitles', () => {
  it('setPrintArea writes _xlnm.Print_Area scoped to the sheet', () => {
    const wb = createWorkbook();
    addWorksheet(wb, 'Report');
    const dn = setPrintArea(wb, 0, 'A1:E20');
    expect(dn.name).toBe('_xlnm.Print_Area');
    // Excel reads an unqualified print area as belonging to whatever sheet is
    // active, so the value has to name the sheet the scope points at.
    expect(dn.value).toBe("'Report'!A1:E20");
    expect(dn.scope).toBe(0);
  });

  it('setPrintArea quotes a sheet title that needs it', () => {
    const wb = createWorkbook();
    addWorksheet(wb, 'Quarter 1');
    expect(setPrintArea(wb, 0, '$A$1:$E$20').value).toBe("'Quarter 1'!$A$1:$E$20");
  });

  it('setPrintArea quotes a title that would otherwise read as a cell or a boolean', () => {
    const wb = createWorkbook();
    addWorksheet(wb, 'A1');
    addWorksheet(wb, 'TRUE');
    expect(setPrintArea(wb, 0, 'A1:E20').value).toBe("'A1'!A1:E20");
    expect(setPrintArea(wb, 1, 'A1:E20').value).toBe("'TRUE'!A1:E20");
  });

  it('setPrintArea doubles an apostrophe in the title', () => {
    const wb = createWorkbook();
    addWorksheet(wb, "Bob's Sheet");
    expect(setPrintArea(wb, 0, 'A1:E20').value).toBe("'Bob''s Sheet'!A1:E20");
  });

  it('setPrintArea leaves an already-qualified ref alone', () => {
    const wb = createWorkbook();
    addWorksheet(wb, 'Report');
    expect(setPrintArea(wb, 0, "'Report'!$A$1:$E$20").value).toBe("'Report'!$A$1:$E$20");
    expect(setPrintArea(wb, 0, 'report!A1:E20').value).toBe('report!A1:E20');
  });

  it('setPrintArea throws when a leg names a different sheet', () => {
    const wb = createWorkbook();
    addWorksheet(wb, 'Report');
    addWorksheet(wb, 'Other');
    expect(() => setPrintArea(wb, 0, 'A1:B2,Other!D1:E2')).toThrow(OpenXmlSchemaError);
  });

  it('setPrintArea qualifies every leg of a multi-area range', () => {
    const wb = createWorkbook();
    addWorksheet(wb, 'Report');
    expect(setPrintArea(wb, 0, 'A1:B2,D1:E2').value).toBe("'Report'!A1:B2,'Report'!D1:E2");
  });

  it('setPrintArea writes a value getDefinedNameTarget can read back', () => {
    const wb = createWorkbook();
    addWorksheet(wb, 'Quarter 1');
    setPrintArea(wb, 0, 'A1:B2,D1:E2');
    const targets = getDefinedNameTarget(wb, '_xlnm.Print_Area', 0);
    expect(targets?.map((t) => t.sheet)).toEqual(['Quarter 1', 'Quarter 1']);
    expect(targets?.map((t) => t.range)).toEqual(['A1:B2', 'D1:E2']);
  });

  it('setPrintArea drops a leading = before qualifying', () => {
    const wb = createWorkbook();
    addWorksheet(wb, 'Report');
    expect(setPrintArea(wb, 0, '=A1:E20').value).toBe("'Report'!A1:E20");
  });

  it('setPrintArea trims every leg, qualified or not', () => {
    const wb = createWorkbook();
    addWorksheet(wb, 'Quarter 1');
    setPrintArea(wb, 0, "A1:B2, 'Quarter 1'!D1:E2");
    const targets = getDefinedNameTarget(wb, '_xlnm.Print_Area', 0);
    expect(targets?.map((t) => t.sheet)).toEqual(['Quarter 1', 'Quarter 1']);
  });

  it('setPrintArea throws on an empty leg', () => {
    const wb = createWorkbook();
    addWorksheet(wb, 'Report');
    expect(() => setPrintArea(wb, 0, 'A1:B2,,D1:E2')).toThrow(OpenXmlSchemaError);
    expect(() => setPrintArea(wb, 0, '')).toThrow(OpenXmlSchemaError);
  });

  it('setPrintArea throws when sheetIndex names no sheet', () => {
    const wb = createWorkbook();
    addWorksheet(wb, 'Report');
    expect(() => setPrintArea(wb, 3, 'A1:E20')).toThrow(OpenXmlSchemaError);
  });

  it('setPrintArea throws when sheetIndex names a chartsheet', () => {
    const wb = createWorkbook();
    addWorksheet(wb, 'Report');
    addChartsheet(wb, 'Chart1');
    expect(() => setPrintArea(wb, 1, 'A1:E20')).toThrow(OpenXmlSchemaError);
  });

  it('setPrintTitles formats both rows + cols with sheet prefix', () => {
    const wb = createWorkbook();
    addWorksheet(wb, 'Report');
    const dn = setPrintTitles(wb, 0, { rows: '$1:$1', cols: '$A:$A', sheetName: 'Report' });
    expect(dn.value).toBe("'Report'!$A:$A,'Report'!$1:$1");
  });

  it('setPrintTitles supports rows-only or cols-only', () => {
    const wb = createWorkbook();
    addWorksheet(wb, 'A');
    const dnRows = setPrintTitles(wb, 0, { rows: '$1:$1', sheetName: 'A' });
    expect(dnRows.value).toBe("'A'!$1:$1");
    const dnCols = setPrintTitles(wb, 0, { cols: '$A:$A', sheetName: 'A' });
    expect(dnCols.value).toBe("'A'!$A:$A");
  });

  it('doubles an apostrophe in the sheet name so the reference stays readable', () => {
    // `Bob's Sheet` is a title `validateSheetTitle` accepts. A raw `'${title}'`
    // produced `'Bob's Sheet'!$1:$1`, which closes the quoted name after
    // `Bob` and leaves `s Sheet'!$1:$1` as garbage.
    const wb = createWorkbook();
    addWorksheet(wb, "Bob's Sheet");
    const dn = setPrintTitles(wb, 0, { rows: '$1:$1', sheetName: "Bob's Sheet" });
    expect(dn.value).toBe("'Bob''s Sheet'!$1:$1");
  });

  it('writes a Print_Titles value getDefinedNameTarget can read back', () => {
    const wb = createWorkbook();
    addWorksheet(wb, "Bob's Sheet");
    setPrintTitles(wb, 0, { rows: '$1:$1', cols: '$A:$A', sheetName: "Bob's Sheet" });
    const targets = getDefinedNameTarget(wb, '_xlnm.Print_Titles', 0);
    expect(targets?.map((t) => t.sheet)).toEqual(["Bob's Sheet", "Bob's Sheet"]);
    expect(targets?.map((t) => t.range)).toEqual(['$A:$A', '$1:$1']);
  });

  it('setPrintTitles throws OpenXmlSchemaError when neither rows nor cols is supplied', () => {
    const wb = createWorkbook();
    addWorksheet(wb, 'A');
    expect(() => setPrintTitles(wb, 0, { sheetName: 'A' })).toThrow(OpenXmlSchemaError);
  });

  it('full save → load round-trip preserves the print-area + titles', async () => {
    const wb = createWorkbook();
    addWorksheet(wb, 'A');
    setPrintArea(wb, 0, 'A1:E20');
    setPrintTitles(wb, 0, { rows: '$1:$1', sheetName: 'A' });

    const bytes = await workbookToBytes(wb);
    const wb2 = await loadWorkbook(fromBuffer(bytes));
    const pa = getDefinedName(wb2, '_xlnm.Print_Area', 0);
    const pt = getDefinedName(wb2, '_xlnm.Print_Titles', 0);
    expect(pa?.value).toBe("'A'!A1:E20");
    expect(pt?.value).toBe("'A'!$1:$1");
  });
});