// Shared by the io tests that assert on the emitted package rather than on the
// reloaded model. Not a suite of its own: vitest only collects `*.test.ts`.

import type { Unzipped } from 'fflate';
import type { Chartsheet } from '../../src/chartsheet/chartsheet.js';
import type { Worksheet } from '../../src/worksheet/worksheet.js';

const decoder = new TextDecoder();

export const SST_PART = 'xl/sharedStrings.xml';
export const SHEET_PART = 'xl/worksheets/sheet1.xml';

/** Narrow a loaded sheet to a worksheet, or fail with what it turned out to be. */
export const expectSheet = (sheet: Worksheet | Chartsheet | undefined): Worksheet => {
  if (!sheet) throw new Error('expected worksheet');
  if (!('rows' in sheet)) throw new Error('expected worksheet, got chartsheet');
  return sheet;
};

/** Decode one part of an unzipped package as text. */
export const partText = (archive: Unzipped, path: string): string => {
  const entry = archive[path];
  if (!entry) throw new Error(`no ${path} in the package`);
  return decoder.decode(entry);
};
