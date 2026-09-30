// Control characters appear as `\uXXXX` escapes rather than literal bytes: a
// literal NUL makes Git classify this file as binary, and every change to it
// then lands as `Bin 2085 -> 3026 bytes` with no reviewable diff.

import { describe, expect, it } from 'vitest';
import { escapeCellString, unescapeCellString } from '../../src/utils/escape.js';

describe('escapeCellString / unescapeCellString', () => {
  it('passes plain ASCII through unchanged', () => {
    const s = 'hello world 123';
    expect(escapeCellString(s)).toBe(s);
    expect(unescapeCellString(s)).toBe(s);
  });

  it('replaces NUL with _x0000_', () => {
    expect(escapeCellString('a\u0000b')).toBe('a_x0000_b');
  });

  it('replaces other illegal control chars with their _xHHHH_ form', () => {
    expect(escapeCellString('\u0001\u0007\u000B\u000C\u001F')).toBe('_x0001__x0007__x000B__x000C__x001F_');
  });

  it('escapes `\\t` / `\\n` / `\\r` so XML CRLF normalisation does not lose them', () => {
    // XML parsers normalise CRLF / lone CR to LF on read. To preserve
    // these whitespace bytes through a load → save → load cycle the
    // writer encodes them via `_xHHHH_`. Decoded form on the way out
    // must match the input.
    const s = '\t\n\r';
    expect(escapeCellString(s)).toBe('_x0009__x000A__x000D_');
    expect(unescapeCellString(escapeCellString(s))).toBe(s);
  });

  it('protects existing _xHHHH_ sequences from being unescaped on round-trip', () => {
    const literal = 'foo_x0041_bar'; // user wrote the literal string foo_x0041_bar
    const esc = escapeCellString(literal);
    expect(esc).toBe('foo_x005F_x0041_bar');
    expect(unescapeCellString(esc)).toBe(literal);
  });

  it('protects adjacent _xHHHH_ sequences that share an underscore', () => {
    // In `_x0041_x0042_` one underscore closes the first sequence and opens
    // the second, and a reader decodes left to right without overlapping, so
    // an opener left unprotected comes back decoded: `_x0041B`.
    const cases: Array<[string, string]> = [
      ['_x0041_x0042_', '_x005F_x0041_x005F_x0042_'],
      ['a_x0044_x0045_x0046_b', 'a_x005F_x0044_x005F_x0045_x005F_x0046_b'],
      ['_x005F_x0041_', '_x005F_x005F_x005F_x0041_'],
      ['__x0041__', '__x005F_x0041__'],
    ];
    for (const [literal, escaped] of cases) {
      expect(escapeCellString(literal)).toBe(escaped);
      expect(unescapeCellString(escaped)).toBe(literal);
    }
  });

  it('protects a sequence whose closing underscore comes from an escaped character', () => {
    // `\n` goes out as `_x000A_`, which puts an underscore straight after
    // `SKU_x0041` and hands the reader a sequence the text never contained.
    const cases: Array<[string, string]> = [
      ['SKU_x0041\nrest', 'SKU_x005F_x0041_x000A_rest'],
      ['col_x0009\tval', 'col_x005F_x0009_x0009_val'],
      ['_x0041\u0001', '_x005F_x0041_x0001_'],
    ];
    for (const [literal, escaped] of cases) {
      expect(escapeCellString(literal)).toBe(escaped);
      expect(unescapeCellString(escaped)).toBe(literal);
    }
  });

  it('leaves an underscore that opens no sequence alone', () => {
    for (const s of ['_x0041', 'x0041_', '_xZZZZ_', '_x041_', '_x00411_']) {
      expect(escapeCellString(s)).toBe(s);
      expect(unescapeCellString(s)).toBe(s);
    }
  });

  it('round-trips a mix of legal and illegal characters', () => {
    const original = 'A\u0000B_x0042_C\u0007';
    const round = unescapeCellString(escapeCellString(original));
    expect(round).toBe(original);
  });

  it('uppercases the hex digits', () => {
    expect(escapeCellString('«')).toBe('«'); // 0xAB is allowed in XML 1.0
    expect(escapeCellString('\u001F')).toBe('_x001F_');
  });

  it('unescapeCellString accepts both upper and lower-case hex (Excel does)', () => {
    expect(unescapeCellString('_x0041_')).toBe('A');
    expect(unescapeCellString('_x0061_')).toBe('a');
    expect(unescapeCellString('_x000a_')).toBe('\n');
  });

  it('leaves astral characters alone', () => {
    // A surrogate pair is one legal XML character. Escaping its halves
    // separately produced `_xD83D__xDE00_`, which Excel shows verbatim.
    for (const s of ['hi \u{1F600} there', '\u{20000}', '\u{1D11E}\u{10437}']) {
      expect(escapeCellString(s)).toBe(s);
      expect(unescapeCellString(escapeCellString(s))).toBe(s);
    }
  });

  it('still escapes an unpaired surrogate, which has no UTF-8 encoding', () => {
    expect(escapeCellString('a\uD83Db')).toBe('a_xD83D_b');
    expect(escapeCellString('a\uDE00b')).toBe('a_xDE00_b');
    expect(escapeCellString('a\uD83D')).toBe('a_xD83D_');
  });

  it('escapes U+FFFE / U+FFFF, which XML 1.0 leaves out of its Char production', () => {
    expect(escapeCellString('a\uFFFEb')).toBe('a_xFFFE_b');
    expect(escapeCellString('a\uFFFFb')).toBe('a_xFFFF_b');
    expect(unescapeCellString(escapeCellString('a\uFFFFb'))).toBe('a\uFFFFb');
    // Only those two BMP codepoints are excluded. The astral noncharacters sit
    // inside `[#x10000-#x10FFFF]`, so they stay verbatim.
    expect(escapeCellString('\u{1FFFE}')).toBe('\u{1FFFE}');
  });

  it('unescapes the surrogate halves older versions of this writer emitted', () => {
    // Files already on disk carry the split form; reading them must still
    // rebuild the character.
    expect(unescapeCellString('hi _xD83D__xDE00_ there')).toBe('hi \u{1F600} there');
  });
});
