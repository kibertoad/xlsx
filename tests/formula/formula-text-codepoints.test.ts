// What XML 1.0 can carry in the text nodes the readers hand back verbatim
// (`<f>`, a `t="str"` cached result, `<formula1>` / `<formula2>`): tab and LF
// as themselves, CR only as a character reference, since a raw CR is
// normalised to LF on read. A C0 control character and an unpaired surrogate
// have no representation at all, and the `_xHHHH_` convention is not one:
// nothing decodes these nodes, so an encoded sequence would reach Excel as the
// literal text `_x0000_`.

import { describe, expect, it } from 'vitest';
import { unzipSync } from 'fflate';
import { isFormulaValue, makeFormula } from '../../src/cell/cell.js';
import { loadWorkbook } from '../../src/io/load.js';
import { fromBuffer } from '../../src/io/node.js';
import { workbookToBytes } from '../../src/io/save.js';
import { OpenXmlSchemaError } from '../../src/utils/exceptions.js';
import { addWorksheet, createWorkbook } from '../../src/workbook/workbook.js';
import { makeDataValidation } from '../../src/worksheet/data-validations.js';
import { addDataValidation, getCell, setCell } from '../../src/worksheet/worksheet.js';

const decoder = new TextDecoder();

const NUL = String.fromCharCode(0x00);
const BEL = String.fromCharCode(0x07);
const LONE_SURROGATE = String.fromCharCode(0xd800);
const EMOJI = String.fromCodePoint(0x1f600);

const save = async (formula: string, cachedValue?: string): Promise<Uint8Array> => {
  const wb = createWorkbook();
  const value = cachedValue === undefined ? makeFormula(formula) : makeFormula(formula, { cachedValue });
  setCell(addWorksheet(wb, 'S'), 1, 1, value);
  return workbookToBytes(wb);
};

const sheetXml = (bytes: Uint8Array): string => {
  const entry = unzipSync(bytes)['xl/worksheets/sheet1.xml'];
  if (!entry) throw new Error('no sheet1.xml in the package');
  return decoder.decode(entry);
};

const reload = async (bytes: Uint8Array): Promise<{ formula: string; cachedValue?: string }> => {
  const wb = await loadWorkbook(fromBuffer(bytes));
  const ref = wb.sheets[0];
  if (ref?.kind !== 'worksheet') throw new Error('expected a worksheet');
  const value = getCell(ref.sheet, 1, 1)?.value;
  if (value === undefined || !isFormulaValue(value)) throw new Error('expected a formula cell');
  const cached = value.cachedValue;
  return { formula: value.formula, ...(typeof cached === 'string' ? { cachedValue: cached } : {}) };
};

describe('codepoints in formula text and cached results', () => {
  it.each(['\uFFFE', '\uFFFF'])('rejects an XML-forbidden noncharacter (%s)', async (text) => {
    await expect(save(`="${text}"`)).rejects.toThrow(OpenXmlSchemaError);
    await expect(save('=A2', text)).rejects.toThrow(OpenXmlSchemaError);
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'S');
    addDataValidation(ws, makeDataValidation({ type: 'custom', sqref: 'A1', formula1: `=A1="${text}"` }));
    await expect(workbookToBytes(wb)).rejects.toThrow(OpenXmlSchemaError);
  });

  it('round-trips CR, LF and tab in formula text', async () => {
    const formula = 'IF(A1,\r\n\t"yes",\n\t"no")';
    const bytes = await save(`=${formula}`);
    // The CR has to travel as a character reference; the parser would collapse
    // a literal one to LF, which is the whole reason for the encoding.
    expect(sheetXml(bytes)).toContain('&#13;');
    expect((await reload(bytes)).formula).toBe(formula);
  });

  it('round-trips CR in a cached string result', async () => {
    const cachedValue = `line1\rline2\r\nline3`;
    const bytes = await save('=A2', cachedValue);
    expect((await reload(bytes)).cachedValue).toBe(cachedValue);
  });

  it('round-trips an astral character in both nodes', async () => {
    const formula = `CONCAT("${EMOJI}")`;
    const bytes = await save(`=${formula}`, EMOJI);
    expect(await reload(bytes)).toEqual({ formula, cachedValue: EMOJI });
  });

  it('rejects a control character in formula text, naming the cell', async () => {
    await expect(save(`=CONCAT("a${BEL}b")`)).rejects.toThrow(
      new OpenXmlSchemaError('worksheet: formula at A1 contains U+0007, which XML 1.0 cannot represent'),
    );
  });

  it('rejects a control character in a cached string result', async () => {
    await expect(save('=A2', `a${NUL}b`)).rejects.toThrow(
      new OpenXmlSchemaError('worksheet: cached formula result at A1 contains U+0000, which XML 1.0 cannot represent'),
    );
  });

  it('rejects an unpaired surrogate in formula text', async () => {
    await expect(save(`=CONCAT("a${LONE_SURROGATE}b")`)).rejects.toThrow(OpenXmlSchemaError);
    await expect(save(`=CONCAT("a${LONE_SURROGATE}b")`)).rejects.toThrow(/U\+D800/);
  });

  it('rejects a control character in a data-validation formula', async () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'S');
    addDataValidation(ws, makeDataValidation({ type: 'custom', sqref: 'A1:A10', formula1: `=A1="a${BEL}b"` }));
    await expect(workbookToBytes(wb)).rejects.toThrow(
      new OpenXmlSchemaError(
        'worksheet: data-validation formula1 at A1:A10 contains U+0007, which XML 1.0 cannot represent',
      ),
    );
  });
});
