// Shared error-message construction for cell text, so that one file is
// rejected in the same words whether the caller loads a workbook, streams it,
// or writes one back out.

import { formatSheetQualifiedRef, tupleToCoordinate } from './coordinate.js';
import { OpenXmlSchemaError } from './exceptions.js';

/** How much of an offending value an error echoes. Cell text is untrusted and unbounded. */
const MAX_QUOTED_LENGTH = 40;

/**
 * How much text Excel accepts in one cell, in UTF-16 code units: the unit
 * `String.length` counts and the one Excel's own `LEN` reports. A longer
 * string is written without complaint by every tool that does not check, and
 * Excel then reports the workbook as needing repair.
 *
 * The unit is code units rather than characters, which decides the astral
 * case: 16_384 emoji are 16_384 characters but 32_768 code units, and Excel
 * refuses that cell the same way it refuses 32_768 letters.
 *
 * Exported so a caller with text of unknown length can shorten it ahead of the
 * write rather than discover the ceiling from a thrown error. A slice must
 * keep a surrogate pair together: splitting one leaves a lone surrogate, which
 * the XML escaper rejects further down with an error about the pair rather
 * than about the length.
 */
export const MAX_CELL_TEXT_LENGTH = 32_767;

/**
 * Reject cell text past Excel's per-cell ceiling. Call it from every
 * `serializeCell` branch that emits text.
 *
 * Takes the length rather than the text so that a rich-text cell can add up
 * its runs instead of joining them, which would allocate a second copy of
 * every such cell on the write path.
 *
 * `what` names the kind of text and `ref` the cell, so a rich-text cell and a
 * cached formula result report the field the caller can actually shorten.
 */
export function requireCellTextFits(length: number, what: string, ref: string): void {
  if (length <= MAX_CELL_TEXT_LENGTH) return;
  throw new OpenXmlSchemaError(
    `worksheet: ${what} at ${ref} is ${length} UTF-16 code units,` +
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
