// A central directory can claim DEFLATE for an entry that carries no
// compressed bytes. No real zip writer does that (even a zero-byte file
// deflates to a final empty block), so it only shows up in a crafted archive.
//
// `readStream` used to spin on it forever: with nothing to push into the
// inflater, `pull` enqueued nothing, closed nothing and errored nothing, so
// the stream machinery called it again with identical state. Reached through
// `loadWorkbookStream`, that is an unkillable busy loop on untrusted input.
// `read` was quieter but no better, handing back an empty part.

import { zipSync } from 'fflate';
import { describe, expect, it } from 'vitest';
import { OpenXmlIoError } from '../../src/utils/exceptions.js';
import { openRandomAccessArchive } from '../../src/zip/random-access-reader.js';

const SIG_LFH = 0x04034b50;
const SIG_CD = 0x02014b50;
const SIG_EOCD = 0x06054b50;
const METHOD_DEFLATE = 8;

/** One-entry archive whose only entry declares DEFLATE with a zero-byte payload. */
const craftEmptyDeflateEntry = (name: string): Uint8Array => {
  const nameBytes = new TextEncoder().encode(name);
  const n = nameBytes.length;

  const lfh = new Uint8Array(30 + n);
  const lfhView = new DataView(lfh.buffer);
  lfhView.setUint32(0, SIG_LFH, true);
  lfhView.setUint16(8, METHOD_DEFLATE, true);
  lfhView.setUint16(26, n, true);
  lfh.set(nameBytes, 30);

  const cd = new Uint8Array(46 + n);
  const cdView = new DataView(cd.buffer);
  cdView.setUint32(0, SIG_CD, true);
  cdView.setUint16(10, METHOD_DEFLATE, true);
  cdView.setUint16(28, n, true);
  cdView.setUint32(42, 0, true);
  cd.set(nameBytes, 46);

  const eocd = new Uint8Array(22);
  const eocdView = new DataView(eocd.buffer);
  eocdView.setUint32(0, SIG_EOCD, true);
  eocdView.setUint16(8, 1, true);
  eocdView.setUint16(10, 1, true);
  eocdView.setUint32(12, cd.length, true);
  eocdView.setUint32(16, lfh.length, true);

  const out = new Uint8Array(lfh.length + cd.length + eocd.length);
  out.set(lfh, 0);
  out.set(cd, lfh.length);
  out.set(eocd, lfh.length + cd.length);
  return out;
};

describe('DEFLATE entry with no compressed bytes', () => {
  const bytes = craftEmptyDeflateEntry('a.xml');

  it('rejects a streaming read instead of spinning', () => {
    const archive = openRandomAccessArchive(bytes);
    expect(() => archive.readStream('a.xml')).toThrow(OpenXmlIoError);
  });

  it('rejects a buffered read rather than returning an empty part', () => {
    const archive = openRandomAccessArchive(bytes);
    expect(() => archive.read('a.xml')).toThrow(OpenXmlIoError);
  });

  it('still streams an entry whose deflate payload decodes to nothing', async () => {
    // The legitimate neighbour of the crafted case: a zero-byte file really
    // does get a deflate stream, just one that emits no bytes. It has to
    // close cleanly, which is the path the guard must not break.
    const empty = zipSync({ 'a.xml': new Uint8Array(0) }, { level: 6 });
    const archive = openRandomAccessArchive(empty);
    const reader = archive.readStream('a.xml').getReader();
    let emitted = 0;
    for (let next = await reader.read(); !next.done; next = await reader.read()) {
      emitted += next.value.byteLength;
    }
    expect(emitted).toBe(0);
    expect(archive.read('a.xml').byteLength).toBe(0);
  });
});
