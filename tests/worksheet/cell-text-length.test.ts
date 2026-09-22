// Excel accepts 32_767 characters in one cell. A longer string used to be
// written without complaint, and Excel then reported the workbook as needing
// repair and truncated the cell. The writer refuses it instead, so the failure
// lands where the caller can still do something about it.

import { describe, expect, it } from 'vitest';
import { makeRichText } from '../../src/cell/rich-text.js';
import { workbookToBytes } from '../../src/io/save.js';
import { createWriteOnlyWorkbook } from '../../src/streaming/write-only.js';
import { MAX_CELL_TEXT_LENGTH } from '../../src/utils/cell-text.js';
import { OpenXmlSchemaError } from '../../src/utils/exceptions.js';
import { addWorksheet, createWorkbook } from '../../src/workbook/workbook.js';
import { setCell } from '../../src/worksheet/worksheet.js';

const AT_LIMIT = 'x'.repeat(MAX_CELL_TEXT_LENGTH);
const OVER_LIMIT = 'x'.repeat(MAX_CELL_TEXT_LENGTH + 1);

const save = async (value: Parameters<typeof setCell>[3]): Promise<Uint8Array> => {
  const wb = createWorkbook();
  setCell(addWorksheet(wb, 'S'), 1, 1, value);
  return workbookToBytes(wb);
};

describe('cell text length', () => {
  it('accepts a string exactly at the limit', async () => {
    await expect(save(AT_LIMIT)).resolves.toBeInstanceOf(Uint8Array);
  });

  it('rejects a string one character past it', async () => {
    await expect(save(OVER_LIMIT)).rejects.toThrow(OpenXmlSchemaError);
  });

  it('names the cell and both lengths', async () => {
    await expect(save(OVER_LIMIT)).rejects.toThrow(/A1 is 32768 characters, past the 32767/);
  });

  it('counts a rich-text cell across its runs, not per run', async () => {
    // Each run is comfortably inside the limit; the cell they make is not.
    const half = 'y'.repeat(MAX_CELL_TEXT_LENGTH / 2 + 1);
    await expect(
      save({ kind: 'rich-text', runs: makeRichText([{ text: half }, { text: half }]) }),
    ).rejects.toThrow(OpenXmlSchemaError);
  });

  it('accepts rich-text runs that together sit at the limit', async () => {
    const half = 'y'.repeat(MAX_CELL_TEXT_LENGTH - 1);
    await expect(
      save({ kind: 'rich-text', runs: makeRichText([{ text: half }, { text: 'z' }]) }),
    ).resolves.toBeInstanceOf(Uint8Array);
  });

  it('counts characters, not UTF-16 code units', async () => {
    // An astral character is two code units but one character, and Excel counts
    // the characters. A `String.length` check would have refused this cell at
    // half the real limit.
    const emoji = '\u{1F600}'.repeat(MAX_CELL_TEXT_LENGTH - 1);
    // Under the limit as characters, well over it as code units.
    expect([...emoji]).toHaveLength(MAX_CELL_TEXT_LENGTH - 1);
    expect(emoji.length).toBeGreaterThan(MAX_CELL_TEXT_LENGTH);
    await expect(save(emoji)).resolves.toBeInstanceOf(Uint8Array);
  });

  it('rejects an over-long astral string too', async () => {
    const emoji = '\u{1F600}'.repeat(MAX_CELL_TEXT_LENGTH + 1);
    await expect(save(emoji)).rejects.toThrow(OpenXmlSchemaError);
  });

  it('rejects an over-long cached formula result', async () => {
    await expect(
      save({ kind: 'formula', t: 'normal', formula: 'A2', cachedValue: OVER_LIMIT }),
    ).rejects.toThrow(OpenXmlSchemaError);
  });

  it('rejects an over-long string on the write-only path as well', async () => {
    const sink = {
      toBytes: () => ({
        write: () => {},
        finish: () => new Uint8Array(0),
      }),
    };
    const wb = await createWriteOnlyWorkbook(sink as never);
    const ws = await wb.addWorksheet('S');
    await expect(ws.appendRow([OVER_LIMIT])).rejects.toThrow(OpenXmlSchemaError);
    wb.abort();
  });
});
