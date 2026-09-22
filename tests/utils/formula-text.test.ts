// OOXML stores formula text without the leading `=` a spreadsheet UI shows
// (ECMA-376 §18.3.1.40). `normalizeFormulaText` is the one place that rule is
// applied, so it has to hold for every spelling a caller can hand it.

import { describe, expect, it } from 'vitest';
import { makeFormula, setFormula, bindValue, makeCell } from '../../src/cell/cell.js';
import { OpenXmlSchemaError } from '../../src/utils/exceptions.js';
import { normalizeFormulaText, startsWithEquals } from '../../src/utils/formula-text.js';

describe('normalizeFormulaText', () => {
  it('drops the UI prefix and the whitespace around it', () => {
    expect(normalizeFormulaText('=SUM(A1:A3)')).toBe('SUM(A1:A3)');
    expect(normalizeFormulaText(' =SUM(A1) ')).toBe('SUM(A1)');
    expect(normalizeFormulaText('= SUM(A1)')).toBe('SUM(A1)');
  });

  it('leaves text that carries no prefix alone', () => {
    expect(normalizeFormulaText('SUM(A1:A3)')).toBe('SUM(A1:A3)');
    expect(normalizeFormulaText('')).toBe('');
  });

  it('keeps an `=` that is part of the expression', () => {
    // A comparison is a formula in its own right; only the first `=` is the
    // UI's, and the rest of the expression is untouched.
    expect(normalizeFormulaText('=A1=B1')).toBe('A1=B1');
    expect(normalizeFormulaText('=IF(A1="x",1,0)')).toBe('IF(A1="x",1,0)');
  });

  it('rejects text that still begins with `=` after the prefix comes off', () => {
    // Stripping again would store `A1` and invent a reading for input nobody
    // meant; keeping the `=` would store `<f>=A1</f>`, which Excel reports as
    // a damaged file. Neither is an answer, so this fails loudly.
    for (const text of ['==A1', '= =A1', '  ==SUM(A1)  ', '===A1']) {
      expect(() => normalizeFormulaText(text)).toThrow(OpenXmlSchemaError);
    }
  });

  it('never returns text a serialiser would have to reject', () => {
    for (const text of ['=A1', 'A1', ' = A1 ', '=-A1', '=A1=B1', '']) {
      expect(startsWithEquals(normalizeFormulaText(text))).toBe(false);
    }
  });
});

describe('the constructors that normalise for the caller', () => {
  it('makeFormula rejects a doubled prefix', () => {
    expect(() => makeFormula('==A1')).toThrow(OpenXmlSchemaError);
  });

  it('setFormula rejects a doubled prefix', () => {
    expect(() => setFormula(makeCell(1, 1), '==A1')).toThrow(OpenXmlSchemaError);
  });

  it('bindValue rejects a doubled prefix rather than writing a damaged cell', () => {
    // `bindValue` routes a leading `=` to the formula path, so this used to
    // land `<f>=A1</f>` on the cell.
    expect(() => bindValue(makeCell(1, 1), '==A1')).toThrow(OpenXmlSchemaError);
  });
});
