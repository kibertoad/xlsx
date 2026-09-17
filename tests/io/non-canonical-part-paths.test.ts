// What binds sharedStrings, styles or the theme to a workbook is its
// relationship; `xl/sharedStrings.xml` and its siblings are where Excel
// happens to write them. So the rels decide what loadWorkbook reads, and a
// rel that resolves to nothing is a malformed package rather than an absent
// part.

import { describe, expect, it } from 'vitest';
import { loadWorkbook } from '../../src/io/load.js';
import { fromBuffer } from '../../src/io/node.js';
import { OpenXmlSchemaError } from '../../src/utils/exceptions.js';
import type { Workbook } from '../../src/workbook/workbook.js';
import { getCell, type Worksheet } from '../../src/worksheet/worksheet.js';
import {
  movePart,
  NUMBER_FORMAT,
  part,
  type Parts,
  relocatedPackage,
  replaceOnce,
  rewrite,
  stalePackage,
  THEME_BYTES,
  WB_RELS_PATH,
  zipParts,
} from './non-canonical-package.js';

const load = (parts: Parts): Promise<Workbook> => loadWorkbook(fromBuffer(zipParts(parts)));

const onlySheet = (wb: Workbook): Worksheet => {
  expect(wb.sheets).toHaveLength(1);
  const ref = wb.sheets[0];
  if (ref?.kind !== 'worksheet') throw new Error(`expected a worksheet, got "${String(ref?.kind)}"`);
  return ref.sheet;
};

const repointSharedStrings = (parts: Parts, attrs: string): void => {
  rewrite(parts, WB_RELS_PATH, (xml) => replaceOnce(xml, 'Target="strings.xml"', attrs));
};

describe('loadWorkbook: workbook-level parts at the path the rels choose', () => {
  it('reads sharedStrings, styles and theme from wherever the rels point', async () => {
    const wb = await load(await relocatedPackage());

    expect(getCell(onlySheet(wb), 1, 1)?.value).toBe('hello');
    expect(wb.styles.cellXfs).toHaveLength(2);
    expect([...wb.styles.numFmts.values()]).toContain(NUMBER_FORMAT);
    expect(wb.themeXml).toEqual(THEME_BYTES);
  });

  it('follows the relationship over a part left behind at the conventional path', async () => {
    const parts = await relocatedPackage();
    const stale = await stalePackage();
    parts['xl/sharedStrings.xml'] = part(stale, 'xl/sharedStrings.xml');
    parts['xl/styles.xml'] = part(stale, 'xl/styles.xml');

    const wb = await load(parts);

    expect(getCell(onlySheet(wb), 1, 1)?.value).toBe('hello');
    expect(wb.styles.cellXfs).toHaveLength(2);
  });

  it('rejects a relationship that targets a part the package does not contain', async () => {
    const parts = await relocatedPackage();
    repointSharedStrings(parts, 'Target="nowhere.xml"');

    await expect(load(parts)).rejects.toThrow(OpenXmlSchemaError);
    await expect(load(parts)).rejects.toThrow(/sharedStrings relationship targets "xl\/nowhere.xml"/);
  });

  it('rejects an external relationship for a workbook-level part', async () => {
    const parts = await relocatedPackage();
    repointSharedStrings(parts, 'Target="https://example.invalid/strings.xml" TargetMode="External"');

    await expect(load(parts)).rejects.toThrow(/sharedStrings relationship is external/);
  });

  it('resolves a percent-encoded target against the literal entry name', async () => {
    const parts = await relocatedPackage();
    movePart(parts, 'xl/strings.xml', 'xl/shared strings.xml');
    repointSharedStrings(parts, 'Target="shared%20strings.xml"');

    const wb = await load(parts);

    expect(getCell(onlySheet(wb), 1, 1)?.value).toBe('hello');
  });
});
