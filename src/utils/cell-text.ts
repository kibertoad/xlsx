// Shared error-message construction for the cell readers, so that one file is
// rejected in the same words whether the caller loads a workbook or streams it.

import { formatSheetQualifiedRef, tupleToCoordinate } from './coordinate.js';
import { OpenXmlSchemaError } from './exceptions.js';

/** How much of an offending value an error echoes. Cell text is untrusted and unbounded. */
const MAX_QUOTED_LENGTH = 40;

/**
 * Characters Excel accepts in one cell's text. A longer string is written
 * without complaint by every tool that does not check, and Excel then reports
 * the workbook as needing repair and truncates the cell.
 *
 * Exported so a caller with text of its own length can truncate or split ahead
 * of the write rather than discover the ceiling from a thrown error.
 */
export const MAX_CELL_TEXT_LENGTH = 32_767;

/**
 * Characters in `text`, counted only far enough to decide whether it is over
 * `MAX_CELL_TEXT_LENGTH`.
 *
 * `String.length` counts UTF-16 code units, which overcounts every astral
 * character by one: 20_000 emoji are 40_000 units but 20_000 characters, and
 * Excel counts the characters. So a string within the limit by units is within
 * it by characters too, and that is the check every ordinary cell takes. Only
 * a string that fails it pays for the codepoint walk, which stops as soon as
 * the answer is known.
 */
const exceedsCellTextLength = (text: string): boolean => {
  if (text.length <= MAX_CELL_TEXT_LENGTH) return false;
  let chars = 0;
  for (const _ of text) {
    chars++;
    if (chars > MAX_CELL_TEXT_LENGTH) return true;
  }
  return false;
};

/**
 * Reject cell text past Excel's per-cell ceiling. `ref` names the cell and
 * `what` the kind of text, so a rich-text cell and a cached formula result
 * report the field the caller can actually shorten.
 */
export function requireCellTextLength(text: string, what: string, ref: string): void {
  if (!exceedsCellTextLength(text)) return;
  throw new OpenXmlSchemaError(
    `worksheet: ${what} at ${ref} is ${[...text].length} characters,` +
      ` past the ${MAX_CELL_TEXT_LENGTH} Excel accepts in one cell`,
  );
}

/** Quote cell text into an error message, bounded, with the full length named. */
export function quoteCellText(raw: string): string {
  if (raw.length <= MAX_QUOTED_LENGTH) return raw;
  return `${raw.slice(0, MAX_QUOTED_LENGTH)}... (${raw.length} chars)`;
}

/** `Sheet1!B2` label for an error message. */
export function cellLabel(sheet: string, col: number, row: number): string {
  return formatSheetQualifiedRef(sheet, tupleToCoordinate(col, row));
}

/** A `t` outside ST_CellType (ECMA-376 §18.18.11). Neither reader can guess what the `<v>` holds. */
export function unknownCellType(t: string, sheet: string, col: number, row: number): OpenXmlSchemaError {
  return new OpenXmlSchemaError(
    `worksheet: unknown cell type t="${quoteCellText(t)}" at ${cellLabel(sheet, col, row)}`,
  );
}
