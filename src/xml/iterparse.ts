// SAX iterator over OOXML XML payloads. Wraps `saxes` (XMLNS-aware) and yields
// a flat stream of {start | end | text} events with names already converted to
// Clark notation (`{ns}local`) — same shape as the DOM parser produces for
// static XmlNode trees, so consumers can switch between bulk and streaming
// reads without retouching name comparison.
//
// Phase 1 §3 acceptance: 1 k–row sheetData walked end-to-end with cell counts
// matching the source. The phase-4 read-only worksheet drives real-world use;
// this layer just produces the events.
//
// DOCTYPE / external entity declarations are forbidden. saxes does not expand
// external entities, but a prescan also rejects DTDs. Every input shape is fed
// to the parser in chunks, and each chunk is prescanned before it is fed, so
// the rejection happens before saxes sees the declaration.

import { SaxesParser } from 'saxes';
import { OpenXmlSchemaError } from '../utils/exceptions.js';
import { qname } from './namespaces.js';

export type SaxEvent =
  | { kind: 'start'; name: string; attrs: Record<string, string> }
  | { kind: 'end'; name: string }
  | { kind: 'text'; text: string };

/**
 * Streamable input: `Uint8Array`, plain string, or a Web `ReadableStream` of
 * `Uint8Array` chunks (produced by xlsx zip entries via fflate, fetch, file
 * streams, etc.).
 */
export type SaxInput = Uint8Array | string | ReadableStream<Uint8Array>;

const DOCTYPE_RE = /<!DOCTYPE\b/;
const ENTITY_RE = /<!ENTITY\b/;

/** Longest token {@link checkDoctype} matches, in code units. */
const DTD_TOKEN_LENGTH = '<!DOCTYPE'.length;

const checkDoctype = (text: string): void => {
  if (DOCTYPE_RE.test(text)) {
    throw new OpenXmlSchemaError('DTD declarations are not permitted in OOXML payloads');
  }
  if (ENTITY_RE.test(text)) {
    throw new OpenXmlSchemaError('Entity declarations are not permitted in OOXML payloads');
  }
};

const isReadableStream = (v: unknown): v is ReadableStream<Uint8Array> => {
  return typeof v === 'object' && v !== null && typeof (v as ReadableStream).getReader === 'function';
};

const decoder = (): TextDecoder => new TextDecoder('utf-8', { fatal: false });

/**
 * Feed size, in code units. saxes runs its handlers synchronously inside
 * `write()`, so one `write()` queues every event its argument produces before
 * the consumer sees any of them. Capping the argument caps the queue.
 *
 * This bounds every input shape, including a `ReadableStream`, whose upstream
 * chunk size is not ours to choose: the zip reader pushes 64 KB of *compressed*
 * bytes per pull, and worksheet XML inflates roughly sevenfold, so its chunks
 * arrive around 450 KB.
 */
const FEED_CHUNK_SIZE = 64 * 1024;

/**
 * Split text that exceeds the feed size. Boundaries fall on arbitrary offsets,
 * which is safe because saxes holds back a lone high surrogate at the end of a
 * `write()` and rejoins it with the next one, so a split codepoint still
 * surfaces as a single text event.
 */
function* atFeedSize(text: string): IterableIterator<string> {
  if (text.length === 0) return;
  if (text.length <= FEED_CHUNK_SIZE) {
    yield text;
    return;
  }
  for (let i = 0; i < text.length; i += FEED_CHUNK_SIZE) {
    yield text.slice(i, i + FEED_CHUNK_SIZE);
  }
}

/** Decode any supported input into a sequence of feed-sized text chunks. */
async function* decodedChunks(input: SaxInput): AsyncIterableIterator<string> {
  if (typeof input === 'string') {
    yield* atFeedSize(input);
    return;
  }
  // `stream: true` holds back a codepoint split across a chunk boundary
  // instead of emitting a replacement character for each half.
  const td = decoder();
  if (input instanceof Uint8Array) {
    for (let i = 0; i < input.byteLength; i += FEED_CHUNK_SIZE) {
      yield* atFeedSize(td.decode(input.subarray(i, Math.min(i + FEED_CHUNK_SIZE, input.byteLength)), { stream: true }));
    }
  } else if (isReadableStream(input)) {
    const reader = input.getReader();
    for (;;) {
      const { done, value } = await reader.read();
      if (done) break;
      yield* atFeedSize(td.decode(value, { stream: true }));
    }
  } else {
    throw new OpenXmlSchemaError('iterParse: unsupported input type');
  }
  yield* atFeedSize(td.decode());
}

interface SaxesOpenTag {
  name: string;
  uri: string;
  local: string;
  prefix: string;
  attributes: Record<string, { value: string; uri: string; local: string; prefix: string }>;
  isSelfClosing?: boolean;
}

interface SaxesCloseTag {
  name: string;
  uri: string;
  local: string;
  prefix: string;
}

const buildAttrsClark = (attrs: SaxesOpenTag['attributes']): Record<string, string> => {
  const out: Record<string, string> = {};
  for (const [, info] of Object.entries(attrs)) {
    // saxes already resolved the namespace when xmlns: true is set; raw xmlns /
    // xmlns:* declarations have prefix='xmlns' (or local==='xmlns' when
    // default) and we drop those — they're rebuilt by the serializer.
    if (info.prefix === 'xmlns' || (info.prefix === '' && info.local === 'xmlns')) continue;
    const key = qname(info.uri, info.local);
    out[key] = info.value;
  }
  return out;
};

/**
 * Parse the input as a stream of SAX events. Element / attribute names are
 * returned in Clark notation (`{ns}local`).
 */
export async function* iterParse(input: SaxInput): AsyncIterableIterator<SaxEvent> {
  // Set up the parser. xmlns: true gives us resolved {uri, local, prefix} on
  // every open / close tag and on every attribute.
  const parser = new SaxesParser({ xmlns: true, fragment: false });

  // Head-pointer ring instead of Array#shift: each saxes write() can produce
  // hundreds of events in a single synchronous batch (a `<row>` with dozens of
  // cells flushes one opentag + one text + one closetag per cell). `shift()`
  // is O(n) per element in V8, so a single-batch drain of N events would be
  // O(N²) before iteration. The head advances on yield; the queue is reset
  // (head + length) once it drains. Peak depth is one chunk's worth of events,
  // which is what keeps the queue bounded on a multi-GB sheet.
  let queue: SaxEvent[] = [];
  let head = 0;
  let pending: Error | undefined;

  parser.on('error', (err: Error) => {
    pending = err;
  });
  parser.on('doctype', () => {
    pending = new OpenXmlSchemaError('DTD declarations are not permitted in OOXML payloads');
  });
  parser.on('opentag', (node: SaxesOpenTag) => {
    queue.push({ kind: 'start', name: qname(node.uri, node.local), attrs: buildAttrsClark(node.attributes) });
  });
  parser.on('closetag', (node: SaxesCloseTag) => {
    queue.push({ kind: 'end', name: qname(node.uri, node.local) });
  });
  parser.on('text', (text: string) => {
    if (text.length > 0) queue.push({ kind: 'text', text });
  });

  const drain = function* (): IterableIterator<SaxEvent> {
    for (;;) {
      const ev = head < queue.length ? queue[head] : undefined;
      if (ev === undefined) {
        // Reset rather than grow forever; the next batch starts at index 0.
        queue = [];
        head = 0;
        return;
      }
      head++;
      yield ev;
    }
  };

  const feed = (chunk: string): void => {
    parser.write(chunk);
    if (pending !== undefined) throw pending;
  };

  // Carry the tail of each chunk into the next scan so a `<!DOCTYPE` straddling
  // a chunk boundary is still matched.
  let dtdCarry = '';
  const scanForDtd = (chunk: string): void => {
    const window = dtdCarry.length === 0 ? chunk : dtdCarry + chunk;
    checkDoctype(window);
    dtdCarry = window.slice(-(DTD_TOKEN_LENGTH - 1));
  };

  for await (const chunk of decodedChunks(input)) {
    scanForDtd(chunk);
    feed(chunk);
    yield* drain();
  }

  parser.close();
  if (pending !== undefined) throw pending;
  yield* drain();
}
