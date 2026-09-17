// `_xNNNN_` in formula text and in a cached string result is literal text.
// Both nodes come back from the reader exactly as stored, with no unescape
// pass, so the writer must not apply the `_xHHHH_` cell-string convention to
// them: an escape nothing reverses drifts further from the caller's text on
// every save.

import { unzipSync } from 'fflate';
import { describe, expect, it } from 'vitest';
import { type FormulaValue, isFormulaValue, makeFormula } from '../../src/cell/cell.js';
import { loadWorkbook } from '../../src/io/load.js';
import { fromBuffer } from '../../src/io/node.js';
import { workbookToBytes } from '../../src/io/save.js';
import { addWorksheet, createWorkbook } from '../../src/workbook/workbook.js';
import { getCell, setCell } from '../../src/worksheet/worksheet.js';

const decoder = new TextDecoder();

const buildWith = async (value: FormulaValue): Promise<Uint8Array> => {
  const wb = createWorkbook();
  setCell(addWorksheet(wb, 'S'), 1, 1, value);
  return workbookToBytes(wb);
};

const formulaOf = async (bytes: Uint8Array): Promise<FormulaValue> => {
  const wb = await loadWorkbook(fromBuffer(bytes));
  const ref = wb.sheets[0];
  if (ref?.kind !== 'worksheet') throw new Error('expected a worksheet');
  const value = getCell(ref.sheet, 1, 1)?.value;
  if (value === undefined || !isFormulaValue(value)) throw new Error('expected a formula cell');
  return value;
};

const sheetXml = (bytes: Uint8Array): string => {
  const entry = unzipSync(bytes)['xl/worksheets/sheet1.xml'];
  if (!entry) throw new Error('no sheet1.xml in the package');
  return decoder.decode(entry);
};

describe('formula text containing a literal _xNNNN_', () => {
  it('is written verbatim', async () => {
    const bytes = await buildWith(makeFormula('=CONCAT("_x0041_")'));
    expect(sheetXml(bytes)).toContain('<f>CONCAT("_x0041_")</f>');
  });

  it('keeps a cached string result verbatim', async () => {
    const bytes = await buildWith(makeFormula('=A2', { cachedValue: '_x0041_' }));
    expect(sheetXml(bytes)).toContain('<v>_x0041_</v>');
    expect((await formulaOf(bytes)).cachedValue).toBe('_x0041_');
  });

  it('is stable across repeated load and save cycles', async () => {
    const formula = 'CONCAT("_x0041_")';
    const cachedValue = '_x0042_';
    let bytes = await buildWith(makeFormula(`=${formula}`, { cachedValue }));
    for (let cycle = 0; cycle < 3; cycle++) {
      bytes = await workbookToBytes(await loadWorkbook(fromBuffer(bytes)));
      const parsed = await formulaOf(bytes);
      expect(parsed.formula).toBe(formula);
      expect(parsed.cachedValue).toBe(cachedValue);
    }
  });

  it('still XML-escapes the characters that need it', async () => {
    const bytes = await buildWith(makeFormula('=IF(A1<2,"a&b","c")', { cachedValue: 'a&b<c' }));
    const xml = sheetXml(bytes);
    expect(xml).toContain('<f>IF(A1&lt;2,"a&amp;b","c")</f>');
    expect(xml).toContain('<v>a&amp;b&lt;c</v>');
    const parsed = await formulaOf(bytes);
    expect(parsed.formula).toBe('IF(A1<2,"a&b","c")');
    expect(parsed.cachedValue).toBe('a&b<c');
  });
});
