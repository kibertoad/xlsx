// ZIP read layer.
//
// `openZip(source)` walks the central directory once and inflates each entry on
// demand inside `read(path)` (see `./random-access-reader.ts`). That keeps peak
// memory at compressed-archive size + per-entry inflate scratch + the bounded
// cache of small re-read entries (`./inflate-cache.ts`), instead of holding
// every uncompressed entry resident at once the way the old `unzipSync`
// shortcut did. The fallback path through fflate's `unzipSync` is preserved for
// ZIP64 / non-standard archives.

import type { XlsxSource } from '../io/source.js';
import { OpenXmlIoError, OpenXmlUnsupportedFormatError, type UnsupportedFormat } from '../utils/exceptions.js';
import { isRawBiffWorkbook } from './biff.js';
import { classifyCfb, isCfbCompoundDocument } from './cfb.js';
import type { DecompressionLimits } from './decompression-guard.js';
import { openRandomAccessArchive } from './random-access-reader.js';

// The two known kinds of CFB input need opposite advice: an encrypted xlsx has
// to be decrypted and a legacy `.xls` has to be converted. Telling the owner of
// an old `.xls` to decrypt it sends them looking for a password that was never
// set. The `.xls` advice still mentions passwords: BIFF encryption lives inside
// the `Workbook` stream, so a protected `.xls` looks like any other from here,
// and the headless conversion fails on it.
const UNSUPPORTED_FORMAT_MESSAGES: Record<UnsupportedFormat, string> = {
  'encrypted-xlsx': 'Encrypted xlsx is not supported. Decrypt with msoffcrypto-tool first.',
  'legacy-xls':
    'This is a legacy .xls (BIFF) workbook, which @office-kit/xlsx does not read. ' +
    'Convert it to .xlsx first: open it in Excel or LibreOffice and save it as .xlsx, or run ' +
    '`soffice --headless --convert-to xlsx`. A password-protected .xls has to be opened with ' +
    'its password and re-saved, which the headless conversion cannot do.',
  'compound-file':
    'The input is an OLE compound file rather than an xlsx package. It is either a ' +
    'password-protected xlsx (decrypt with msoffcrypto-tool first) or a legacy Office ' +
    'format such as .xls (convert it to .xlsx first).',
};

const unsupportedFormat = (format: UnsupportedFormat): OpenXmlUnsupportedFormatError =>
  new OpenXmlUnsupportedFormatError(format, UNSUPPORTED_FORMAT_MESSAGES[format]);

export interface ZipArchive {
  /** Sorted list of all entry paths in the archive. */
  list(): string[];
  /**
   * Synchronous read; throws OpenXmlIoError when the path is unknown. Each
   * call returns an array the caller owns: mutating it changes neither the
   * archive nor what a later read of the same path returns.
   */
  read(path: string): Uint8Array;
  /** Promise variant for symmetry with the future streaming reader. */
  readAsync(path: string): Promise<Uint8Array>;
  /**
   * Streaming read: returns the entry's inflated bytes as a Web
   * `ReadableStream<Uint8Array>` chunk-by-chunk. Lets callers (the streaming
   * worksheet iterator, in particular) push the inflated payload through a SAX
   * parser without first materialising it in full — peak memory for a sheet
   * walk drops to the inflate window + SAX state instead of the entire
   * uncompressed worksheet body. Throws OpenXmlIoError when the path is
   * unknown.
   */
  readStream(path: string): ReadableStream<Uint8Array>;
  /** Whether the archive holds an entry at the given path. */
  has(path: string): boolean;
  /** Release the in-memory entry table. Subsequent reads throw. */
  close(): void;
}

/** Options for {@link openZip}. */
export interface OpenZipOptions {
  /**
   * Decompression-bomb safeguards applied while inflating archive entries. The
   * default limits admit any legitimate xlsx and reject pathological archives
   * (extreme compression ratios, gigabyte-scale entries). Pass `false` to
   * disable the guard entirely — only safe when the source is fully trusted.
   * See {@link DecompressionLimits} for the individual knobs.
   */
  decompressionLimits?: DecompressionLimits | false;
}

/**
 * Open a zip archive from any {@link XlsxSource}. The source is fully
 * materialised in memory, the central directory is parsed once, and each
 * entry is inflated on demand by {@link openRandomAccessArchive}: peak memory
 * stays at compressed-archive size, plus per-entry inflate scratch, plus a few
 * MB at most of small entries kept for re-reads, rather than holding every
 * uncompressed entry resident. The fflate `unzipSync` fallback is preserved
 * internally for some malformed ZIP32 archives the random-access reader
 * rejects, and it does hold every entry inflated.
 *
 * Throws {@link OpenXmlIoError} when the bytes are not a readable zip, and
 * when the source itself fails to produce them. See the error contract on
 * {@link OpenXmlIoError}. An OLE compound file or a raw BIFF workbook throws
 * {@link OpenXmlUnsupportedFormatError} instead, carrying the format it
 * recognised.
 */
export async function openZip(source: XlsxSource, opts: OpenZipOptions = {}): Promise<ZipArchive> {
  let bytes: Uint8Array;
  try {
    bytes = await source.toBytes();
  } catch (cause) {
    throw new OpenXmlIoError('openZip: failed to read source bytes', { cause });
  }

  // Neither an encrypted xlsx nor a legacy `.xls` is a zip: both are OLE
  // Compound File Binary containers. Reject them here with advice that fits
  // the one at hand rather than letting the zip reader fail with a generic
  // invalid-archive message.
  if (isCfbCompoundDocument(bytes)) throw unsupportedFormat(classifyCfb(bytes));
  // A `.xls` from Excel 4.0 or earlier has no compound file around it.
  if (isRawBiffWorkbook(bytes)) throw unsupportedFormat('legacy-xls');

  return openRandomAccessArchive(bytes, opts.decompressionLimits);
}
