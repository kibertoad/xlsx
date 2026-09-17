// `setHyperlinks` / `setComments` exist to avoid a scan per entry. What they
// must not change is the result, so these compare a batch against the loop of
// single calls it stands in for, including the two orderings the single calls
// have: setHyperlink moves a replaced entry to the end, setComment replaces it
// where it sits.

import { describe, expect, it } from 'vitest';
import { makeHyperlink } from '../../src/worksheet/hyperlinks.js';
import { makeLegacyComment } from '../../src/worksheet/comments.js';
import {
  makeWorksheet,
  setComment,
  setComments,
  setHyperlink,
  setHyperlinks,
  type Worksheet,
} from '../../src/worksheet/worksheet.js';

type LinkEntry = { ref: string; target: string };

const linkEntries = (refs: ReadonlyArray<string>): LinkEntry[] =>
  refs.map((ref, i) => ({ ref, target: `https://e.example/${ref}/${i}` }));

const viaLoop = (ws: Worksheet, entries: ReadonlyArray<LinkEntry>): void => {
  for (const entry of entries) setHyperlink(ws, entry.ref, { target: entry.target });
};

describe('setHyperlinks matches the loop it replaces', () => {
  it('appends fresh refs in order', () => {
    const entries = linkEntries(['A1', 'A2', 'A3']);
    const batched = makeWorksheet('S');
    const looped = makeWorksheet('S');

    setHyperlinks(batched, entries);
    viaLoop(looped, entries);

    expect(batched.hyperlinks).toEqual(looped.hyperlinks);
    expect(batched.hyperlinks.map((h) => h.ref)).toEqual(['A1', 'A2', 'A3']);
  });

  it('moves a replaced entry to the end, as a single call does', () => {
    const entries = [...linkEntries(['A1', 'A2']), { ref: 'A1', target: 'https://again.example' }];
    const batched = makeWorksheet('S');
    const looped = makeWorksheet('S');

    setHyperlinks(batched, entries);
    viaLoop(looped, entries);

    expect(batched.hyperlinks).toEqual(looped.hyperlinks);
    expect(batched.hyperlinks.map((h) => h.ref)).toEqual(['A2', 'A1']);
  });

  it('consumes duplicate refs one at a time, as a run of single calls does', () => {
    // Duplicates are reachable by pushing onto the array, and a file can
    // declare two entries for one cell. Each set replaces the first of them.
    const seed = (ws: Worksheet): void => {
      ws.hyperlinks.push(makeHyperlink({ ref: 'A1', target: 'https://first.example' }));
      ws.hyperlinks.push(makeHyperlink({ ref: 'A1', target: 'https://second.example' }));
      ws.hyperlinks.push(makeHyperlink({ ref: 'B2', target: 'https://b.example' }));
    };
    const entries = [
      { ref: 'A1', target: 'https://third.example' },
      { ref: 'A1', target: 'https://fourth.example' },
    ];

    const batched = makeWorksheet('S');
    const looped = makeWorksheet('S');
    seed(batched);
    seed(looped);

    setHyperlinks(batched, entries);
    viaLoop(looped, entries);

    expect(batched.hyperlinks).toEqual(looped.hyperlinks);
  });

  it('keeps the array identity, so a held reference sees the result', () => {
    const ws = makeWorksheet('S');
    const held = ws.hyperlinks;
    setHyperlinks(ws, linkEntries(['A1', 'A2']));
    expect(held).toBe(ws.hyperlinks);
    expect(held).toHaveLength(2);
  });

  it('writes nothing when an entry is invalid', () => {
    const ws = makeWorksheet('S');
    setHyperlinks(ws, linkEntries(['A1']));

    expect(() =>
      setHyperlinks(ws, [
        { ref: 'A2', target: 'https://ok.example' },
        { ref: 'A3' },
        { ref: 'A4', target: 'https://ok.example' },
      ]),
    ).toThrow(/entry 1 \("A3"\)/);
    expect(ws.hyperlinks.map((h) => h.ref)).toEqual(['A1']);
  });

  it('resolves a ref without scanning the entries already there', () => {
    // Counting `ref` reads measures the thing that made the loop quadratic
    // without depending on wall-clock: a scan reads every entry's ref, the
    // batch index reads each once.
    const counter = { reads: 0 };
    const N = 5000;
    const ws = makeWorksheet('S');
    for (let i = 0; i < N; i++) {
      const ref = `A${i}`;
      ws.hyperlinks.push({
        get ref() {
          counter.reads++;
          return ref;
        },
        target: 'https://e.example',
      });
    }

    counter.reads = 0;
    setHyperlinks(ws, linkEntries(Array.from({ length: N }, (_, i) => `B${i}`)));
    // One read per entry already on the sheet, to index it. The loop of single
    // calls reads N per call.
    expect(counter.reads).toBeLessThan(2 * N);
    expect(ws.hyperlinks).toHaveLength(2 * N);
  });
});

describe('setComments matches the loop it replaces', () => {
  const commentEntries = (refs: ReadonlyArray<string>) =>
    refs.map((ref, i) => ({ ref, author: 'a', text: `note ${ref} ${i}` }));

  it('replaces an existing ref where it sits, as a single call does', () => {
    const entries = [...commentEntries(['A1', 'A2']), { ref: 'A1', author: 'b', text: 'edited' }];
    const batched = makeWorksheet('S');
    const looped = makeWorksheet('S');

    setComments(batched, entries);
    for (const entry of entries) setComment(looped, entry);

    expect(batched.legacyComments).toEqual(looped.legacyComments);
    expect(batched.legacyComments.map((c) => c.ref)).toEqual(['A1', 'A2']);
    expect(batched.legacyComments[0]?.text).toBe('edited');
  });

  it('leaves a duplicate ref alone past the first, as a single call does', () => {
    const seed = (ws: Worksheet): void => {
      ws.legacyComments.push(makeLegacyComment({ ref: 'A1', author: 'a', text: 'first' }));
      ws.legacyComments.push(makeLegacyComment({ ref: 'A1', author: 'a', text: 'second' }));
    };
    const entries = commentEntries(['A1', 'A1']);

    const batched = makeWorksheet('S');
    const looped = makeWorksheet('S');
    seed(batched);
    seed(looped);

    setComments(batched, entries);
    for (const entry of entries) setComment(looped, entry);

    expect(batched.legacyComments).toEqual(looped.legacyComments);
  });
});
