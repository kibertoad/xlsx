// The batch setters are only worth having if they cannot diverge from the
// single calls they stand in for. The interleavings that break a batch are the
// ones nobody writes by hand, so they are generated: entries seeded straight
// onto the array (duplicates included), then a batch against the same sequence
// applied one call at a time, compared entry by entry.

import fc from 'fast-check';
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

const REFS = ['A1', 'A2', 'A3', 'B1'] as const;
const refArb = fc.constantFrom(...REFS);
const seedArb = fc.array(refArb, { maxLength: 6 });
const batchArb = fc.array(refArb, { maxLength: 12 });

describe('a batch is indistinguishable from the calls it replaces', () => {
  it('leaves hyperlinks exactly as a loop of setHyperlink would', () => {
    fc.assert(
      fc.property(seedArb, batchArb, (seeded, batch) => {
        const entries = batch.map((ref, i) => ({ ref, target: `https://e.example/${ref}/${i}` }));
        const seed = (ws: Worksheet): void => {
          for (const [i, ref] of seeded.entries()) {
            ws.hyperlinks.push(makeHyperlink({ ref, target: `https://seed.example/${ref}/${i}` }));
          }
        };

        const batched = makeWorksheet('S');
        const looped = makeWorksheet('S');
        seed(batched);
        seed(looped);

        setHyperlinks(batched, entries);
        for (const entry of entries) setHyperlink(looped, entry.ref, { target: entry.target });

        expect(batched.hyperlinks).toEqual(looped.hyperlinks);
      }),
    );
  });

  it('leaves comments exactly as a loop of setComment would', () => {
    fc.assert(
      fc.property(seedArb, batchArb, (seeded, batch) => {
        const entries = batch.map((ref, i) => ({ ref, author: `a${i}`, text: `note ${ref} ${i}` }));
        const seed = (ws: Worksheet): void => {
          for (const [i, ref] of seeded.entries()) {
            ws.legacyComments.push(makeLegacyComment({ ref, author: 'seed', text: `seed ${ref} ${i}` }));
          }
        };

        const batched = makeWorksheet('S');
        const looped = makeWorksheet('S');
        seed(batched);
        seed(looped);

        setComments(batched, entries);
        for (const entry of entries) setComment(looped, entry);

        expect(batched.legacyComments).toEqual(looped.legacyComments);
      }),
    );
  });
});
