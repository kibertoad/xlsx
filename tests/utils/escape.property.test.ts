// The invariant the examples in `escape.test.ts` sample: escaping a string and
// unescaping the result returns the string, whatever it held. The inputs that
// break it are built from the convention's own characters, which a general
// string generator produces about never, so the first alphabet is stacked with
// underscores, hex digits and characters the writer has to escape in turn.

import fc from 'fast-check';
import { describe, expect, it } from 'vitest';
import { escapeCellString, unescapeCellString } from '../../src/utils/escape.js';

const conventionChars = fc.constantFrom(
  '_',
  'x',
  '0',
  '1',
  '4',
  '5',
  'F',
  'a',
  '\n',
  '\r',
  '\t',
  '\u0001',
  '￿',
  '\uD83D',
  '\u{1F600}',
);

describe('escapeCellString / unescapeCellString round-trip', () => {
  it('holds for strings made of underscores, hex digits and escaped characters', () => {
    fc.assert(
      fc.property(fc.string({ unit: conventionChars, maxLength: 24 }), (s) => {
        expect(unescapeCellString(escapeCellString(s))).toBe(s);
      }),
    );
  });

  it('holds for arbitrary text, including lone surrogates', () => {
    fc.assert(
      fc.property(fc.string({ unit: 'binary', maxLength: 64 }), (s) => {
        expect(unescapeCellString(escapeCellString(s))).toBe(s);
      }),
    );
  });
});
