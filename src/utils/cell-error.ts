// Error token of a `t="e"` cell, shared by the two worksheet readers so that
// one file is answered the same way whether the caller loads a workbook or
// streams it.

import { cellLabel, quoteCellText } from './cell-text.js';
import { OpenXmlSchemaError } from './exceptions.js';

/**
 * Shape every error token Excel has written shares: the `#` sigil, a letter,
 * then the upper-case ASCII, digits, `_` and the `/` of `#DIV/0!`, closed by
 * the `!` or `?` that all but `#N/A` carry.
 *
 * The shape is what lets an unlisted token be kept: `serializeCell` puts the
 * token straight into `<v>`, and a `<v>` holding markup, a quote or a NUL
 * would leave the saved part malformed or carry a cell of the attacker's
 * choosing into it. The length bound is loose (the longest token in use is
 * `#GETTING_DATA`) and keeps a megabyte of `<v>` text out of the cell model.
 */
const ERROR_TOKEN = /^#[A-Z][A-Z0-9_/]{0,29}[!?]?$/;

/** True for a token a `t="e"` cell can hold, listed in `ERROR_CODES` or not. */
export const isExcelErrorToken = (raw: string): raw is `#${string}` => ERROR_TOKEN.test(raw);

/**
 * Token of an error cell, or `null` when the `<v>` is absent or blank, which
 * is an empty cell the way it is under `t="n"` and `t="b"`.
 *
 * `t="e"` is the file declaring that the cell holds an error, so the token is
 * kept verbatim even when `ERROR_CODES` does not list it: Excel keeps adding
 * tokens, and dropping one loses the cell's value to buy nothing. A `t="e"`
 * carrying something that is not a token at all is a producer bug, and is
 * reported rather than stored.
 *
 * The return type spells the sigil out instead of naming `ExcelErrorCode`
 * because nothing under `src/utils/` depends on the cell model.
 */
export function parseCellErrorCode(
  raw: string | undefined,
  sheet: string,
  col: number,
  row: number,
): `#${string}` | null {
  // Both readers keep element text verbatim, so a `<v>` in a pretty-printed
  // part arrives with its indentation attached.
  const text = raw?.trim();
  if (text === undefined || text === '') return null;
  if (isExcelErrorToken(text)) return text;
  throw new OpenXmlSchemaError(
    `worksheet: <v>${quoteCellText(text)}</v> at ${cellLabel(sheet, col, row)} is not an Excel error token`,
  );
}
