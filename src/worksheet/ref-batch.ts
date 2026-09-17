// Batch application for the worksheet's ref-keyed side tables (hyperlinks,
// legacy comments).
//
// Setting one entry has to find the entry that already carries its ref, which
// is a scan of the array. Paying that per entry is quadratic in the number of
// entries, and a sheet with a link or a note on every row hits it directly.
//
// A batch resolves every ref against one index, built for the batch and dropped
// with it. Nothing survives the call, so the arrays stay what they are: plain,
// public, and safe to edit from anywhere else.

/** What an addition does to the entry already carrying its ref. */
export type RefBatchMode =
  /** Replaces it where it sits, as setComment does. */
  | 'in-place'
  /** Removes it and appends the addition, as setHyperlink does. */
  | 'move-to-end';

interface RefPositions {
  slots: number[];
  head: number;
}

/**
 * Every position each ref occupies, ascending. "The first entry carrying this
 * ref" is then the head of its list rather than a scan.
 *
 * The lists hold every position, not just the first, because these arrays can
 * carry two entries for one cell (a malformed file declares them) and a run of
 * single calls consumes those one at a time.
 */
const positionsByRef = (items: ReadonlyArray<{ ref: string } | undefined>): Map<string, RefPositions> => {
  const positions = new Map<string, RefPositions>();
  for (let i = 0; i < items.length; i++) {
    const item = items[i];
    if (!item) continue;
    const ref = item.ref;
    const known = positions.get(ref);
    if (known) known.slots.push(i);
    else positions.set(ref, { slots: [i], head: 0 });
  }
  return positions;
};

/**
 * Apply `additions` to `items` in order, with the result of setting each one on
 * its own, in one pass over each instead of a scan per addition.
 *
 * `items` keeps its identity, so a caller holding the array sees the result.
 */
export function applyRefBatch<T extends { ref: string }>(
  items: T[],
  additions: ReadonlyArray<T>,
  mode: RefBatchMode,
): void {
  if (additions.length === 0) return;
  const positions = positionsByRef(items);

  if (mode === 'in-place') {
    for (const addition of additions) {
      const ref = addition.ref;
      const at = positions.get(ref)?.slots[0];
      if (at === undefined) {
        items.push(addition);
        positions.set(ref, { slots: [items.length - 1], head: 0 });
      } else {
        items[at] = addition;
      }
    }
    return;
  }

  // A removal tombstones its slot rather than splicing, which is what keeps the
  // pass linear. Survivors keep their relative order, which is what a splice
  // leaves behind too, so the result matches the one-at-a-time sequence.
  const slots: Array<T | undefined> = [...items];
  for (const addition of additions) {
    const ref = addition.ref;
    const known = positions.get(ref);
    // Advancing a cursor avoids shifting every remaining duplicate position.
    const replaced = known?.slots[known.head++];
    if (replaced !== undefined) slots[replaced] = undefined;
    slots.push(addition);
    if (known) known.slots.push(slots.length - 1);
    else positions.set(ref, { slots: [slots.length - 1], head: 0 });
  }

  items.length = 0;
  for (const slot of slots) {
    if (slot) items.push(slot);
  }
}
