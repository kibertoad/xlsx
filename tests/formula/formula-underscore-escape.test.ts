// `<f>` text is a plain XML text node. The writer used to run it through
// `escapeCellString` as well, which rewrites a literal `_xNNNN_` to
// `_x005F_xNNNN_`. Nothing unescapes `<f>` on read, so the rewrite was applied
// again on the next save and the formula drifted further from what the caller
// wrote each time. The same applied to a `t="str"` cached result.

import { unzipSync } from 'fflate';
import { describe, expect, it } from 'vitest';
import { type FormulaValue, makeFormula } from '../../src/cell/cell.js';
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
  const sheet = wb.sheets[0]?.sheet;
  if (!sheet || !('rows' in sheet)) throw new Error('expected a worksheet');
  const value = getCell(sheet, 1, 1)?.value;
  if (value === null || typeof value !== 'object' || !('kind' in value) || value.kind !== 'formula') {
    throw new Error('expected a formula cell');
  }
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

  it('is stable across repeated load and save cycles', async () => {
    const source = 'CONCAT("_x0041_")';
    let bytes = await buildWith(makeFormula(source));
    for (let pass = 0; pass < 3; pass++) {
      expect((await formulaOf(bytes)).formula).toBe(source);
      bytes = await workbookToBytes(await loadWorkbook(fromBuffer(bytes)));
    }
  });

  it('keeps a cached string result verbatim', async () => {
    const bytes = await buildWith(makeFormula('=A2', { cachedValue: '_x0041_' }));
    expect(sheetXml(bytes)).toContain('<v>_x0041_</v>');
    expect((await formulaOf(bytes)).cachedValue).toBe('_x0041_');
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
