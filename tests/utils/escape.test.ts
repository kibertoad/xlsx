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
    // `_x0041_x0042_` is two sequences whose second one starts on the first
    // one's trailing underscore. Protecting only the first left `x0042_`
    // exposed, and the read side decoded it: the cell came back `_x0041B`.
    const literal = '_x0041_x0042_';
    const esc = escapeCellString(literal);
    expect(esc).toBe('_x005F_x0041_x005F_x0042_');
    expect(unescapeCellString(esc)).toBe(literal);
  });

  it('protects a literal _x005F_ that opens another sequence', () => {
    const literal = '_x005F_x0041_';
    expect(unescapeCellString(escapeCellString(literal))).toBe(literal);
  });

  it('round-trips a run of three sequences sharing underscores', () => {
    const literal = 'a_x0044_x0045_x0046_b';
    expect(unescapeCellString(escapeCellString(literal))).toBe(literal);
  });

  it('leaves an underscore that opens no sequence alone', () => {
    for (const s of ['_x0041', 'x0041_', '__x0041__', '_xZZZZ_', '_x041_']) {
      expect(escapeCellString(s)).toBe(s.replace('_x0041_', '_x005F_x0041_'));
      expect(unescapeCellString(escapeCellString(s))).toBe(s);
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
