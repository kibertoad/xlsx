// loadWorkbookStream has to agree with loadWorkbook on two things: where a
// workbook-level part lives, which the rels decide, and which packages are
// malformed. A disagreement surfaces as an empty string table or a missing
// sheet name, neither of which a caller can tell from a workbook that
// genuinely has no shared strings or no such sheet.

import { describe, expect, it } from 'vitest';
import { loadWorkbook } from '../../src/io/load.js';
import { fromBuffer } from '../../src/io/node.js';
import { loadWorkbookStream, type ReadOnlyWorkbook } from '../../src/streaming/read-only.js';
import { OpenXmlSchemaError } from '../../src/utils/exceptions.js';
import {
  part,
  type Parts,
  relocatedPackage,
  replaceOnce,
  rewrite,
  stalePackage,
  WB_RELS_PATH,
  WORKBOOK_PATH,
  zipParts,
} from '../io/non-canonical-package.js';

const open = (parts: Parts): Promise<ReadOnlyWorkbook> => loadWorkbookStream(fromBuffer(zipParts(parts)));

const valuesOf = async (wb: ReadOnlyWorkbook, name: string): Promise<unknown[][]> => {
  const rows: unknown[][] = [];
  for await (const row of wb.openWorksheet(name).iterValues()) rows.push(row);
  return rows;
};

/** Both loaders have to reject the package, and for the same stated reason. */
const bothReject = async (parts: Parts, reason: RegExp): Promise<void> => {
  await expect(open(parts)).rejects.toThrow(OpenXmlSchemaError);
  await expect(open(parts)).rejects.toThrow(reason);
  await expect(loadWorkbook(fromBuffer(zipParts(parts)))).rejects.toThrow(reason);
};

describe('loadWorkbookStream: parity with loadWorkbook on part and sheet resolution', () => {
  it('reads sharedStrings and styles from wherever the rels point', async () => {
    const wb = await open(await relocatedPackage());
    try {
      expect(await valuesOf(wb, 'S')).toEqual([['hello'], [42]]);
      expect(wb.styles.cellXfs).toHaveLength(2);
    } finally {
      await wb.close();
    }
  });

  it('follows the relationship over a part left behind at the conventional path', async () => {
    const parts = await relocatedPackage();
    const stale = await stalePackage();
    parts['xl/sharedStrings.xml'] = part(stale, 'xl/sharedStrings.xml');
    parts['xl/styles.xml'] = part(stale, 'xl/styles.xml');

    const wb = await open(parts);
    try {
      expect(await valuesOf(wb, 'S')).toEqual([['hello'], [42]]);
      expect(wb.styles.cellXfs).toHaveLength(2);
    } finally {
      await wb.close();
    }
  });

  it('rejects a sheet whose r:id has no rels entry instead of dropping the sheet', async () => {
    const parts = await relocatedPackage();
    rewrite(parts, WORKBOOK_PATH, (xml) => replaceOnce(xml, 'r:id="rId1"', 'r:id="rId99"'));

    await bothReject(parts, /sheet "S" rId "rId99" has no matching rels entry/);
  });

  it('rejects a duplicate sheet name instead of keeping the last part under it', async () => {
    const parts = await relocatedPackage();
    rewrite(parts, WORKBOOK_PATH, (xml) =>
      replaceOnce(xml, '</sheets>', '<sheet name="S" sheetId="2" r:id="rId1"/></sheets>'),
    );

    await bothReject(parts, /duplicate sheet name "S"/);
  });

  it('rejects a workbook that declares sheets with no rels part instead of listing none', async () => {
    const parts = await relocatedPackage();
    delete parts[WB_RELS_PATH];

    await bothReject(parts, /workbook has sheets but rels part "xl\/_rels\/workbook\.xml\.rels" is missing/);
  });
});
