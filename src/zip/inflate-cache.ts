// Bounded cache of inflated ZIP entries.
//
// A load comes back to a handful of parts more than once (the `.rels` files the
// relationship walk revisits, above all), and inflating those again is waste.
// Keeping every entry is worse than that waste: it puts the whole uncompressed
// package back in memory, which is the cost the random-access reader exists to
// avoid. So the cache takes small entries only, holds a few MB of them at most,
// and drops the least recently used first. Anything it turns away or evicts
// inflates again on the next read.

/**
 * Largest entry the cache accepts. Sized for the parts a load revisits, which
 * are small; a sheet or a media blob is normally read once, so keeping one buys
 * nothing and costs its full inflated size for the life of the archive.
 */
export const CACHE_MAX_ENTRY_BYTES = 64 * 1024;

/**
 * Ceiling on everything the cache holds, so an archive with thousands of small
 * parts cannot accumulate without limit.
 */
export const CACHE_MAX_TOTAL_BYTES = 4 * 1024 * 1024;

export interface InflateCache {
  /**
   * Cached bytes for `path`, or `undefined`. The caller owns the array it gets
   * back and may mutate it. A hit counts as a use for eviction order.
   */
  get(path: string): Uint8Array | undefined;
  /** Offer `bytes` for later re-reads. Entries above the size ceiling are ignored. */
  set(path: string, bytes: Uint8Array): void;
  clear(): void;
}

export function createInflateCache(): InflateCache {
  // Insertion-ordered, so the first key `keys()` yields is the least recently
  // used for as long as every hit re-inserts.
  const held = new Map<string, Uint8Array>();
  let total = 0;

  return {
    get(path) {
      const hit = held.get(path);
      if (!hit) return undefined;
      // `set` on an existing key leaves it where it was, so a hit has to delete
      // first. Without this the entries being re-read are the first evicted.
      held.delete(path);
      held.set(path, hit);
      return hit.slice();
    },
    set(path, bytes) {
      if (bytes.byteLength > CACHE_MAX_ENTRY_BYTES) return;
      total += bytes.byteLength - (held.get(path)?.byteLength ?? 0);
      // Hold a copy, so the array the caller was handed stays theirs to mutate.
      held.set(path, bytes.slice());
      for (const [lru, entry] of held) {
        if (total <= CACHE_MAX_TOTAL_BYTES) break;
        held.delete(lru);
        total -= entry.byteLength;
      }
    },
    clear() {
      held.clear();
      total = 0;
    },
  };
}
