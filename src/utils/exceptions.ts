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

/**
 * Base class for errors reported by this library, including invalid input,
 * invalid API arguments, unsupported features and configured resource limits.
 *
 * For loads, retrying unchanged bytes with unchanged options does not resolve
 * parsing or validation errors. Source I/O failures may be transient; inspect
 * their cause and the source's retry semantics. A limit error does not establish
 * that the rest of the workbook is valid.
 *
 * Unexpected native errors from library internals are worth reporting, but an
 * error outside this hierarchy can also originate in caller-provided code.
 *
 * Which subclass arrives says where the input broke, not how badly, and the
 * subclass is stable for a given kind of damage. Message text is not: it names
 * parts, offsets and cell references to make a failure diagnosable, so it
 * changes freely between releases. Branch on the class, not on the message.
 */
export class OpenXmlError extends Error {
  override readonly name: string = 'OpenXmlError';

  constructor(message: string, options?: OpenXmlErrorOptions) {
    super(message, options as ErrorOptions);
  }
}

/**
 * Thrown for ZIP, file system, network or stream-level failures: the bytes
 * never arrived, or they did and are not a readable zip archive.
 *
 * This is the class a file that is not an xlsx at all lands on, because the
 * failure happens before any OOXML is parsed. Where a magic number identifies
 * the input (a PDF, a byte-order mark, a zip with no central directory) the
 * message names it, so a CSV renamed to `.xlsx` says so rather than only
 * "not a valid zip".
 *
 * Source read failures preserve the fs / fetch / stream error as `cause` and
 * may be transient. Recreate a consumed source before retrying when necessary.
 * Corrupt archive bytes require a corrected file. This class is also used by
 * write APIs and is the parent of {@link OpenXmlDecompressionBombError}, so
 * the class alone does not determine whether retrying can help.
 */
export class OpenXmlIoError extends OpenXmlError {
  override readonly name: string = 'OpenXmlIoError';
}

/**
 * Thrown when an OOXML payload violates structural / schema invariants: the
 * archive opened, and a part inside it does not parse or contradicts the spec
 * (unreadable XML, a missing required relationship, a cell whose declared type
 * does not match its value).
 *
 * Also used for invalid load options and API arguments. The same class guards
 * the write-side model, where it reports
 * the calling code's mistake rather than a file's: a duplicate sheet title, a
 * merge overlapping an existing one, a style id belonging to another
 * workbook's pool.
 */
export class OpenXmlSchemaError extends OpenXmlError {
  override readonly name = 'OpenXmlSchemaError';
}

/**
 * Thrown when a workbook is structurally valid OOXML but semantically broken.
 *
 * No path in the library throws this today; the semantic checks that exist all
 * report {@link OpenXmlSchemaError}. Catching it is therefore a dead branch.
 */
export class OpenXmlInvalidWorkbookError extends OpenXmlError {
  override readonly name = 'OpenXmlInvalidWorkbookError';
}

/**
 * Thrown for features the port has chosen not to implement (yet), including
 * input that is a real Office format this library does not read, such as an
 * encrypted xlsx or a legacy `.xls`.
 *
 * A supported representation is needed, for example by decrypting the file
 * or re-saving it as `.xlsx` before loading.
 */
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
 *
 * A cap can also reject a legitimate large workbook. Review the input and
 * resource budget before changing the limits; the error does not establish
 * whether the file is valid or hostile.
 */
export class OpenXmlDecompressionBombError extends OpenXmlIoError {
  override readonly name = 'OpenXmlDecompressionBombError';
}

/**
 * Thrown when a read exceeds the cell or row cap configured through
 * `contentLimits` on {@link loadWorkbook} / {@link loadWorkbookStream}. The
 * read stopped at the limit, so the remaining content is not validated. Extends
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
