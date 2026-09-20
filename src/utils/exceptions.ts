// Error hierarchy for @office-kit/xlsx. Public APIs throw subclasses of OpenXmlError;
// internals chain via the `cause` option (Node 18+ / modern browsers all
// support Error.cause).
//
// Class is the one explicitly allowed exception to the no-class rule because
// Error subclasses are how `instanceof` discrimination is expressed in JS.

export interface OpenXmlErrorOptions {
  /** Underlying cause; preserved on the standard `cause` property. */
  cause?: unknown;
}

export class OpenXmlError extends Error {
  override readonly name: string = 'OpenXmlError';

  constructor(message: string, options?: OpenXmlErrorOptions) {
    super(message, options as ErrorOptions);
  }
}

/** Thrown for ZIP, file system, network or stream-level failures. */
export class OpenXmlIoError extends OpenXmlError {
  override readonly name: string = 'OpenXmlIoError';
}

/** Thrown when an OOXML payload violates structural / schema invariants. */
export class OpenXmlSchemaError extends OpenXmlError {
  override readonly name = 'OpenXmlSchemaError';
}

/** Thrown when a workbook is structurally valid OOXML but semantically broken. */
export class OpenXmlInvalidWorkbookError extends OpenXmlError {
  override readonly name = 'OpenXmlInvalidWorkbookError';
}

/** Thrown for features the port has chosen not to implement (yet). */
export class OpenXmlNotImplementedError extends OpenXmlError {
  override readonly name: string = 'OpenXmlNotImplementedError';
}

/**
 * What an input turned out to be when it is a file format this library does
 * not read: `encrypted-xlsx` is a password-protected xlsx, `legacy-xls` a BIFF
 * `.xls`, and `compound-file` an OLE compound file that is neither (another
 * legacy Office format, or a file too damaged to tell).
 */
export type UnsupportedFormat = 'encrypted-xlsx' | 'legacy-xls' | 'compound-file';

/**
 * Thrown by {@link openZip} / {@link loadWorkbook} / {@link loadWorkbookStream}
 * when the input is a recognised file format other than an xlsx package.
 * Subclass of {@link OpenXmlNotImplementedError} so existing
 * `catch (OpenXmlNotImplementedError)` paths still see it. Branch on
 * {@link OpenXmlUnsupportedFormatError.format} to answer "ask for the password"
 * differently from "ask for a re-save"; the message wording is free to change.
 */
export class OpenXmlUnsupportedFormatError extends OpenXmlNotImplementedError {
  override readonly name = 'OpenXmlUnsupportedFormatError';
  readonly format: UnsupportedFormat;

  constructor(format: UnsupportedFormat, message: string, options?: OpenXmlErrorOptions) {
    super(message, options);
    this.format = format;
  }
}

/**
 * Thrown when an archive trips the decompression-bomb safeguards configured on
 * {@link openZip} / {@link loadWorkbook} / {@link loadWorkbookStream}. Subclass
 * of {@link OpenXmlIoError} so existing `catch (OpenXmlIoError)` paths still
 * see it, while letting callers branch on bomb-specific recovery (reject the
 * upload, log a security event, etc.).
 */
export class OpenXmlDecompressionBombError extends OpenXmlIoError {
  override readonly name = 'OpenXmlDecompressionBombError';
}

/**
 * Thrown when a read exceeds the cell or row cap configured through
 * `contentLimits` on {@link loadWorkbook} / {@link loadWorkbookStream}. The
 * workbook is valid; it is larger than the caller allowed. Extends
 * {@link OpenXmlError} directly rather than {@link OpenXmlIoError}, since
 * nothing failed at the I/O layer and a caller wants to tell "too big" apart
 * from "corrupt" to answer an upload with the right status.
 */
export class OpenXmlContentLimitError extends OpenXmlError {
  override readonly name = 'OpenXmlContentLimitError';
}

/**
 * Name an argument's type for an error message, without putting the value in
 * it. Internal (not re-exported from `@office-kit/xlsx/utils`): it exists so a
 * public function handed the wrong argument from JS says what it received
 * instead of interpolating an object and reporting `[object Object]`.
 */
export function describeArg(v: unknown): string {
  if (v === null) return 'null';
  if (v === undefined) return 'undefined';
  if (Array.isArray(v)) return 'an array';
  const t = typeof v;
  return t === 'object' ? 'an object' : `a ${t}`;
}
