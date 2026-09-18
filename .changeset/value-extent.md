---
"@office-kit/xlsx": minor
---

feat: `getValueExtent`, so a sheet formatted past its data can be iterated without rows of `null`

A cell that exists only to carry a style still counts towards the used range,
which is the right answer for Excel and the `<dimension>` element but means a
sheet with formatting on 200 rows and values in 4 iterates 200 rows, 196 of
them entirely `null`. `getValueExtent` returns the bounding box of the cells
holding a value (neither `null` nor `''`), and its shape is accepted directly
by `iterRows` / `iterValues`:

```ts
const box = getValueExtent(ws);
if (box) for (const row of iterValues(ws, box)) { /* 4 rows */ }
```

**Breaking:** `getDataExtent` is renamed to `getCellExtent`, since "data
extent" read like the new value extent while it counts every materialised
cell. Behaviour is unchanged; update the import and the call.

Also newly exported: `getSheetByIndex` from `@office-kit/xlsx/workbook`, which
bounds-checks the index and returns `undefined` for a chartsheet tab, and
`getCellExtentRef` from `@office-kit/xlsx/worksheet`, the plain `"A1:C5"` form
`makeAutoFilter` and `makeTableDefinition` take for their `ref`.
