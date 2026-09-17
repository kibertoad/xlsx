import { readFileSync } from 'node:fs';
import { dirname, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import { fromBuffer } from '../../src/io/node.js';
import { OpenXmlSchemaError } from '../../src/utils/exceptions.js';
import { iterParse, type SaxEvent } from '../../src/xml/iterparse.js';
import { REL_NS, SHEET_MAIN_NS } from '../../src/xml/namespaces.js';
import { openZip } from '../../src/zip/reader.js';

const here = dirname(fileURLToPath(import.meta.url));
const FIXTURES = resolve(here, '../../reference/openpyxl/openpyxl/tests/data/genuine');

const collect = async (input: Parameters<typeof iterParse>[0]): Promise<SaxEvent[]> => {
  const events: SaxEvent[] = [];
  for await (const e of iterParse(input)) events.push(e);
  return events;
};

describe('iterParse — basics', () => {
  it('emits start / end events for a self-closing element', async () => {
    expect(await collect('<r/>')).toEqual([
      { kind: 'start', name: 'r', attrs: {} },
      { kind: 'end', name: 'r' },
    ]);
  });

  it('emits a text event between start and end for a text-only element', async () => {
    expect(await collect('<t>hello</t>')).toEqual([
      { kind: 'start', name: 't', attrs: {} },
      { kind: 'text', text: 'hello' },
      { kind: 'end', name: 't' },
    ]);
  });

  it('preserves attribute values as Clark-keyed strings', async () => {
    const xml = `<workbook xmlns="${SHEET_MAIN_NS}" xmlns:r="${REL_NS}"><sheet r:id="rId1" name="S1"/></workbook>`;
    const events = await collect(xml);
    const start = events.find((e) => e.kind === 'start' && e.name.endsWith('}sheet'));
    expect(start).toBeDefined();
    expect(start?.kind).toBe('start');
    if (start?.kind !== 'start') return;
    expect(start.name).toBe(`{${SHEET_MAIN_NS}}sheet`);
    expect(start.attrs).toEqual({ name: 'S1', [`{${REL_NS}}id`]: 'rId1' });
  });

  it('drops xmlns / xmlns:* declarations from event attributes', async () => {
    const events = await collect(`<r xmlns="${SHEET_MAIN_NS}" xmlns:r="${REL_NS}" foo="bar"/>`);
    const start = events.find((e) => e.kind === 'start');
    expect(start?.kind).toBe('start');
    if (start?.kind !== 'start') return;
    expect(Object.keys(start.attrs).sort()).toEqual(['foo']);
  });

  it('accepts a Uint8Array input', async () => {
    const bytes = new TextEncoder().encode('<r a="1"/>');
    const events = await collect(bytes);
    expect(events.length).toBe(2);
  });

  it('expands the standard XML entities in text', async () => {
    const events = await collect('<t>&amp; &lt; &gt;</t>');
    const text = events.find((e) => e.kind === 'text');
    expect(text?.kind === 'text' && text.text).toBe('& < >');
  });
});

describe('iterParse — security', () => {
  it('rejects DOCTYPE declarations', async () => {
    await expect(collect('<!DOCTYPE foo SYSTEM "u"><foo/>')).rejects.toBeInstanceOf(OpenXmlSchemaError);
  });

  it('rejects loose <!ENTITY declarations', async () => {
    await expect(collect('<!ENTITY foo "bar"><r/>')).rejects.toBeInstanceOf(OpenXmlSchemaError);
  });
});

describe('iterParse — streaming input', () => {
  it('consumes a ReadableStream of Uint8Array chunks', async () => {
    const xml = `<workbook xmlns="${SHEET_MAIN_NS}"><sheets><sheet name="A"/><sheet name="B"/></sheets></workbook>`;
    const enc = new TextEncoder();
    const stream = new ReadableStream<Uint8Array>({
      start(controller) {
        // Split deliberately mid-tag to exercise the chunked path.
        controller.enqueue(enc.encode(xml.slice(0, 100)));
        controller.enqueue(enc.encode(xml.slice(100)));
        controller.close();
      },
    });
    const events = await collect(stream);
    const sheetStarts = events.filter((e) => e.kind === 'start' && e.name === `{${SHEET_MAIN_NS}}sheet`);
    expect(sheetStarts.length).toBe(2);
  });

  it('rejects DOCTYPE hidden in the stream prologue', async () => {
    const enc = new TextEncoder();
    const stream = new ReadableStream<Uint8Array>({
      start(controller) {
        controller.enqueue(enc.encode('<!DOCTYPE x><x/>'));
        controller.close();
      },
    });
    await expect(collect(stream)).rejects.toBeInstanceOf(OpenXmlSchemaError);
  });
});

describe('iterParse — openpyxl genuine/sample.xlsx sheet1.xml', () => {
  it('counts every <row> and <c> element matching the file', async () => {
    const zip = await openZip(fromBuffer(readFileSync(resolve(FIXTURES, 'sample.xlsx'))));
    const events = await collect(zip.read('xl/worksheets/sheet1.xml'));

    let rows = 0;
    let cells = 0;
    for (const e of events) {
      if (e.kind === 'start' && e.name === `{${SHEET_MAIN_NS}}row`) rows++;
      if (e.kind === 'start' && e.name === `{${SHEET_MAIN_NS}}c`) cells++;
    }
    // Cross-checked against grep output captured this turn (2 rows, 2 cells).
    expect(rows).toBe(2);
    expect(cells).toBe(2);
  });

  it('start/end events nest correctly (matched count, balanced order)', async () => {
    const zip = await openZip(fromBuffer(readFileSync(resolve(FIXTURES, 'sample.xlsx'))));
    const events = await collect(zip.read('xl/worksheets/sheet1.xml'));
    const stack: string[] = [];
    let starts = 0;
    let ends = 0;
    for (const e of events) {
      if (e.kind === 'start') {
        starts++;
        stack.push(e.name);
      } else if (e.kind === 'end') {
        ends++;
        const top = stack.pop();
        expect(top).toBe(e.name);
      }
    }
    expect(stack).toEqual([]);
    expect(starts).toBe(ends);
    expect(starts).toBeGreaterThan(0);
  });
});

describe('iterParse: chunked feeding', () => {
  const repeat = (n: number, body: (i: number) => string): string =>
    Array.from({ length: n }, (_, i) => body(i)).join('');

  /**
   * Join adjacent text events. A chunk boundary inside a text node splits it
   * across two `write()` calls and saxes reports one event per call, so this
   * is the granularity at which chunked and unchunked parsing have to agree.
   */
  const joinText = (events: SaxEvent[]): SaxEvent[] => {
    const out: SaxEvent[] = [];
    for (const e of events) {
      const prev = out[out.length - 1];
      if (e.kind === 'text' && prev?.kind === 'text') out[out.length - 1] = { kind: 'text', text: prev.text + e.text };
      else out.push(e);
    }
    return out;
  };

  it('yields events before it has parsed the tail of the document', async () => {
    // saxes reports the mismatched close tag part-way through. Written in one
    // call, every event ahead of it is queued behind the throw and the consumer
    // sees none of them; fed in chunks, the leading elements come out first.
    const xml = `<root>${repeat(8000, (i) => `<c r="A${i}"><v>${i}</v></c>`)}</wrong>`;
    expect(xml.length).toBeGreaterThan(64 * 1024);

    let seen = 0;
    await expect(
      (async () => {
        for await (const _e of iterParse(xml)) seen++;
      })(),
    ).rejects.toThrow(OpenXmlSchemaError);
    expect(seen).toBeGreaterThan(0);
  });

  it('yields events before it has parsed the tail of an oversized stream chunk', async () => {
    // A stream's chunk size is the producer's choice, not ours: the zip reader
    // pushes 64 KB of compressed bytes per pull and worksheet XML inflates
    // about sevenfold, so its chunks arrive around 450 KB. Feeding one of those
    // to saxes whole queues everything it contains, which is the same failure
    // as feeding a whole document, reached through the stream input instead.
    const xml = `<root>${repeat(20_000, (i) => `<c r="A${i}"><v>${i}</v></c>`)}</wrong>`;
    const bytes = new TextEncoder().encode(xml);
    // Comfortably past the 450 KB a real inflated worksheet chunk runs to.
    expect(bytes.byteLength).toBeGreaterThan(450 * 1024);

    const oneChunk = new ReadableStream<Uint8Array>({
      start(controller) {
        controller.enqueue(bytes);
        controller.close();
      },
    });

    let seen = 0;
    await expect(
      (async () => {
        for await (const _e of iterParse(oneChunk)) seen++;
      })(),
    ).rejects.toThrow(OpenXmlSchemaError);
    expect(seen).toBeGreaterThan(0);
  });

  it('produces identical events for string, byte and stream input', async () => {
    // Multi-byte and surrogate-pair characters straddle the chunk boundaries,
    // which is where a naive slice splits a codepoint in half.
    const xml = `<root>${repeat(4000, (i) => `<t k="\u{1F600}${i}">日本語-${i}-\u{1F680}</t>`)}</root>`;
    expect(xml.length).toBeGreaterThan(64 * 1024);
    const bytes = new TextEncoder().encode(xml);

    const fromString = await collect(xml);
    const fromBytes = await collect(bytes);
    const fromStream = await collect(
      new ReadableStream<Uint8Array>({
        start(controller) {
          // 1000-byte chunks land mid-codepoint far more often than 64 KB ones.
          for (let i = 0; i < bytes.length; i += 1000) controller.enqueue(bytes.subarray(i, i + 1000));
          controller.close();
        },
      }),
    );

    expect(fromBytes).toEqual(fromString);
    expect(fromStream).toEqual(fromString);
    const text = fromString.filter((e) => e.kind === 'text').map((e) => (e.kind === 'text' ? e.text : ''));
    expect(text.join('')).toContain('日本語-3999-\u{1F680}');
  });

  it('keeps a surrogate pair intact when it lands exactly on a chunk boundary', async () => {
    // Chunked feeding splits at arbitrary offsets and relies on saxes holding
    // back a lone high surrogate until the next write. Pin the pair across the
    // boundary so that dependency fails loudly if saxes ever stops doing it.
    const CHUNK = 64 * 1024;
    const open = '<r><t>';
    const text = `${'a'.repeat(CHUNK - 1 - open.length)}\u{1F600}b`;
    const xml = `${open}${text}</t></r>`;
    expect(xml.indexOf('\u{1F600}')).toBe(CHUNK - 1);

    const joined = (await collect(xml))
      .filter((e) => e.kind === 'text')
      .map((e) => (e.kind === 'text' ? e.text : ''))
      .join('');
    expect(joined).toBe(text);
    expect([...joined].at(-2)).toBe('\u{1F600}');
  });

  it('rejects an <!ENTITY declaration that straddles a feed-chunk boundary', async () => {
    // A loose `<!ENTITY` is caught by the prescan alone: saxes reports it as a
    // plain syntax error, so only the prescan turns it into an
    // OpenXmlSchemaError. Landing the token across the 64 KB boundary means no
    // single chunk contains all of it, which is what the carry is for.
    const CHUNK = 64 * 1024;
    // Four characters of the token fall in the first chunk, the rest in the next.
    const comment = `<!--${' '.repeat(CHUNK - 4 - '<!---->'.length)}-->`;
    expect(comment.length).toBe(CHUNK - 4);

    const err = await collect(`${comment}<!ENTITY foo "bar"><x/>`).then(
      () => undefined,
      (e: unknown) => e,
    );
    expect(err).toBeInstanceOf(OpenXmlSchemaError);
    expect((err as Error).message).toContain('Entity declarations');
  });

  it('matches a hand-derived event list for a document larger than one chunk', async () => {
    // Comparing the three input shapes against each other only proves they
    // agree: all of them run through the same chunker, so a chunker that
    // dropped or reordered events would corrupt all three identically. This
    // pins the absolute expectation the chunking has to preserve.
    const n = 4000;
    const xml = `<root>${repeat(n, (i) => `<t k="${i}">text-${i}</t>`)}</root>`;
    expect(xml.length).toBeGreaterThan(64 * 1024);

    const expected: SaxEvent[] = [{ kind: 'start', name: 'root', attrs: {} }];
    for (let i = 0; i < n; i++) {
      expected.push({ kind: 'start', name: 't', attrs: { k: String(i) } });
      expected.push({ kind: 'text', text: `text-${i}` });
      expected.push({ kind: 'end', name: 't' });
    }
    expected.push({ kind: 'end', name: 'root' });

    expect(joinText(await collect(xml))).toEqual(expected);
  });

  it('cancels the source stream when the consumer stops early', async () => {
    // Abandoning the iteration is routine: `iterRows({ maxRow })` returns the
    // moment the band ends. The zip reader drops its inflate state from
    // `cancel()`, so a reader that is merely dereferenced keeps the whole
    // decompression window alive for as long as the archive is open.
    let cancelled = false;
    const encoder = new TextEncoder();
    const stream = new ReadableStream<Uint8Array>({
      start(controller) {
        controller.enqueue(encoder.encode('<root>'));
      },
      pull(controller) {
        controller.enqueue(encoder.encode('<c/>'.repeat(1000)));
      },
      cancel() {
        cancelled = true;
      },
    });

    let seen = 0;
    for await (const _e of iterParse(stream)) {
      if (++seen > 3) break;
    }
    expect(cancelled).toBe(true);
  });

  it('reports a syntax error as OpenXmlSchemaError with the saxes error as cause', async () => {
    const err = await collect('<root><a></b></root>').then(
      () => undefined,
      (e: unknown) => e,
    );
    expect(err).toBeInstanceOf(OpenXmlSchemaError);
    expect((err as Error).cause).toBeInstanceOf(Error);
  });
});
