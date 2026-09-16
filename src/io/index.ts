// I/O surface: byte-level Source / Sink types, browser-safe byte helpers
// (Blob / Response / ReadableStream / ArrayBuffer adapters), and the
// xlsx load / save / serialise entry points.

export type { BufferedSinkWriter, XlsxSink } from './sink.js';
export type { XlsxSource } from './source.js';
export {
  fromArrayBuffer,
  fromBlob,
  fromResponse,
  fromStream,
  toArrayBuffer,
  toBlob,
} from './browser.js';
export type { LoadOptions } from './load.js';
export { loadWorkbook } from './load.js';
export type { SaveOptions } from './save.js';
export { saveWorkbook, workbookToBytes } from './save.js';
// Lives in ./zip (the writer owns it), re-exported here because it is the type
// of a SaveOptions field: a caller holding the level in a variable should not
// have to reach into an escape-hatch subpath to name it.
export type { CompressionLevel } from '../zip/writer.js';
