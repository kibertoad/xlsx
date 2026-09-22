// Formula-text normalisation, shared by every path that carries a formula.
//
// A spreadsheet UI shows `=SUM(A1:A3)`, but OOXML stores the text without the
// leading `=` (ECMA-376 §18.3.1.40) in `<f>`, `<formula1>`, `<cfRule>`'s
// `<formula>` and `<definedName>` alike, and Excel reports a file that carries
// one as damaged. The `make*` constructors normalise what a caller hands them,
// and the serialisers normalise again: the model interfaces are public, so a
// caller can build one as a literal and reach the writer without passing a
// constructor.

import { OpenXmlSchemaError } from './exceptions.js';

const EQUALS_SIGN_CODE = 61;

/** True iff `text` is spelled the way a spreadsheet UI shows a formula. */
export function startsWithEquals(text: string): boolean {
  return text.charCodeAt(0) === EQUALS_SIGN_CODE;
}

/**
 * Drop the leading `=` OOXML formula text must not carry, and the whitespace
 * around it: `' =SUM(A1)'` would otherwise keep the `=` in the stored text.
 *
 * Exactly one `=` comes off, because that is the one the UI spelling adds. Text
 * that still starts with `=` after it is gone (`'==A1'`, `'= =A1'`) is not a
 * formula anyone meant, and stripping again would invent a reading for it. It
 * throws {@link OpenXmlSchemaError} instead: the alternative is storing `<f>=A1</f>`,
 * the shape this function exists to keep out of the file.
 */
export function normalizeFormulaText(text: string): string {
  const trimmed = text.trim();
  if (!startsWithEquals(trimmed)) return trimmed;
  const body = trimmed.slice(1).trim();
  if (startsWithEquals(body)) {
    throw new OpenXmlSchemaError(
      `formula text ${JSON.stringify(text)} still begins with "=" once the leading one is removed;` +
        ' OOXML stores formula text without it and Excel reports a file that carries one as damaged',
    );
  }
  return body;
}
