// The `t` attribute on a formula cell types its cached result, and the reader
// and writer have to agree on it. An error result written as `t="str"` is
// ordinary text until Excel recalculates: ISERROR / IFERROR stop matching it
// and an error-keyed conditional format stops firing. A result held in the sst
// (`t="s"`) stores an index, not text, so a reader that keeps the `<v>`
// verbatim hands the index to the caller and the writer then displays the
// number.

import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
import { describe, expect, it } from 'vitest';
import { isFormulaValue, makeFormula } from '../../src/cell/cell.js';
import { makeRichText, makeTextRun } from '../../src/cell/rich-text.js';
import { loadWorkbook } from '../../src/io/load.js';
import { fromBuffer } from '../../src/io/node.js';
import { workbookToBytes } from '../../src/io/save.js';
import type { CellValue } from '../../src/cell/cell.js';
import { addWorksheet, createWorkbook } from '../../src/workbook/workbook.js';
import { getCell, setCell } from '../../src/worksheet/worksheet.js';

const SHEET_PART = 'xl/worksheets/sheet1.xml';

const build = async (value: CellValue): Promise<Uint8Array> => {
  const wb = createWorkbook();
  setCell(addWorksheet(wb, 'S'), 1, 1, value);
  return workbookToBytes(wb);
};

const sheetXml = (bytes: Uint8Array): string => {
  const entry = unzipSync(bytes)[SHEET_PART];
  if (!entry) throw new Error(`no ${SHEET_PART} in the package`);
  return strFromU8(entry);
};

/**
 * Give A1 a formula while leaving its `t` and `<v>` alone. Nothing in the
 * public API writes a formula whose cached result lives in the sst, so the
 * fixture for that shape has to be assembled by hand.
 */
const addFormulaToA1 = (bytes: Uint8Array, formula: string): Uint8Array => {
  const files = unzipSync(bytes);
  const xml = sheetXml(bytes);
  const patched = xml.replace(/(<c r="A1"[^>]*>)(<v>)/, `$1<f>${formula}</f>$2`);
  if (patched === xml) throw new Error(`A1 not found in ${SHEET_PART}: ${xml}`);
  files[SHEET_PART] = strToU8(patched);
  return zipSync(files);
};

const cachedValueOf = async (bytes: Uint8Array): Promise<number | string | boolean | undefined> => {
  const wb = await loadWorkbook(fromBuffer(bytes));
  const ref = wb.sheets[0];
  if (ref?.kind !== 'worksheet') throw new Error('expected a worksheet');
  const value = getCell(ref.sheet, 1, 1)?.value;
  if (value === undefined || !isFormulaValue(value)) throw new Error('expected a formula cell');
  return value.cachedValue;
};

const resave = async (bytes: Uint8Array): Promise<Uint8Array> =>
  workbookToBytes(await loadWorkbook(fromBuffer(bytes)));

describe('cached formula result of error type', () => {
  it('is written as t="e"', async () => {
    const xml = sheetXml(await build(makeFormula('=1/0', { cachedValue: '#DIV/0!' })));
    expect(xml).toContain('<c r="A1" t="e"><f>1/0</f><v>#DIV/0!</v></c>');
  });

  it('keeps its type across a load and save cycle', async () => {
    const bytes = await resave(await build(makeFormula('=NA()', { cachedValue: '#N/A' })));
    expect(await cachedValueOf(bytes)).toBe('#N/A');
    expect(sheetXml(bytes)).toContain('t="e"');
  });

  it('leaves a plain string result on t="str"', async () => {
    const xml = sheetXml(await build(makeFormula('=A2', { cachedValue: 'not an error' })));
    expect(xml).toContain('<c r="A1" t="str"><f>A2</f><v>not an error</v></c>');
  });
});

describe('cached formula result held in the shared-strings table', () => {
  it('resolves to its text rather than the index', async () => {
    const bytes = addFormulaToA1(await build('hello world'), 'INDEX(Data,1)');
    expect(await cachedValueOf(bytes)).toBe('hello world');
  });

  it('is re-saved as t="str" carrying the text', async () => {
    const bytes = await resave(addFormulaToA1(await build('hello world'), 'INDEX(Data,1)'));
    expect(sheetXml(bytes)).toContain('<c r="A1" t="str"><f>INDEX(Data,1)</f><v>hello world</v></c>');
  });

  it('flattens a rich-text entry, which a formula result cannot carry', async () => {
    const runs = makeRichText([makeTextRun('bold', { b: true }), makeTextRun(' plain')]);
    const bytes = addFormulaToA1(await build({ kind: 'rich-text', runs }), 'INDEX(Data,1)');
    expect(await cachedValueOf(bytes)).toBe('bold plain');
  });
});
