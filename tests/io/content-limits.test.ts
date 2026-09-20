// `contentLimits` bounds the quantity that decides what a load costs, which is
// cells, not inflated bytes. A service accepting uploads from strangers has to
// be able to say "nothing over N cells" and have the load stop there rather
// than after the model is built.

import { zipSync } from 'fflate';
import { describe, expect, it } from 'vitest';
import { loadWorkbook } from '../../src/io/load.js';
import { fromBuffer, toBuffer } from '../../src/io/node.js';
import { workbookToBytes } from '../../src/io/save.js';
import { loadWorkbookStream, type ReadOnlyWorksheet } from '../../src/streaming/read-only.js';
import { OpenXmlContentLimitError, OpenXmlError } from '../../src/utils/exceptions.js';
import { createWriteOnlyWorkbook } from '../../src/streaming/write-only.js';
import { addWorksheet, createWorkbook, getSheet } from '../../src/workbook/workbook.js';
import { getCell, setCell } from '../../src/worksheet/worksheet.js';

/** `rows` x `cols` of numbers on one sheet named Data. */
const archive = async (rows: number, cols: number): Promise<Uint8Array> => {
  const sink = toBuffer();
  const wb = await createWriteOnlyWorkbook(sink);
  const ws = await wb.addWorksheet('Data');
  for (let r = 0; r < rows; r++) {
    const row = new Array<number>(cols);
    for (let c = 0; c < cols; c++) row[c] = r * cols + c;
    await ws.appendRow(row);
  }
  await ws.close();
  await wb.finalize();
  return sink.result();
};

/** Two sheets of `rows` x 1, so a cap can be shown to cover the workbook. */
const twoSheetArchive = async (rows: number): Promise<Uint8Array> => {
  const wb = createWorkbook();
  for (const title of ['First', 'Second']) {
    const ws = addWorksheet(wb, title);
    for (let r = 1; r <= rows; r++) setCell(ws, r, 1, r);
  }
  return workbookToBytes(wb);
};

const drain = async (ws: ReadOnlyWorksheet, minRow?: number): Promise<unknown[][]> => {
  const rows: unknown[][] = [];
  for await (const row of ws.iterValues(minRow === undefined ? {} : { minRow })) rows.push(row);
  return rows;
};

describe('loadWorkbook with no contentLimits', () => {
  it('reads the workbook, as it always did', async () => {
    const wb = await loadWorkbook(fromBuffer(await archive(20, 4)));
    const ws = getSheet(wb, 'Data');
    if (ws === undefined) throw new Error('load produced no worksheet');
    expect(getCell(ws, 20, 4)?.value).toBe(79);
  });
});

describe('loadWorkbook with a cell cap', () => {
  it('reads a workbook that fits', async () => {
    const bytes = await archive(10, 4);
    const wb = await loadWorkbook(fromBuffer(bytes), { contentLimits: { maxCells: 40 } });
    expect(getSheet(wb, 'Data')).toBeDefined();
  });

  it('refuses one that does not, naming the cap and the cell that reached it', async () => {
    const bytes = await archive(10, 4);
    await expect(loadWorkbook(fromBuffer(bytes), { contentLimits: { maxCells: 39 } })).rejects.toThrow(
      OpenXmlContentLimitError,
    );
    await expect(loadWorkbook(fromBuffer(bytes), { contentLimits: { maxCells: 39 } })).rejects.toThrow(
      'worksheet: reading Data!D10 passes contentLimits.maxCells of 39',
    );
  });

  it('counts across every worksheet, not per sheet', async () => {
    const bytes = await twoSheetArchive(30);
    // 30 cells on each of two sheets: a 40-cell cap fits the first sheet and
    // has to stop partway through the second.
    await expect(loadWorkbook(fromBuffer(bytes), { contentLimits: { maxCells: 40 } })).rejects.toThrow(
      'worksheet: reading Second!A11 passes contentLimits.maxCells of 40',
    );
    await expect(loadWorkbook(fromBuffer(bytes), { contentLimits: { maxCells: 60 } })).resolves.toBeDefined();
  });

  it('stops before the cells past the cap are modelled', async () => {
    // 200k cells would take seconds and hundreds of MB to model. A cap of 10
    // has to refuse in the time it takes to read ten cells, which is what
    // makes the option worth having.
    const bytes = await archive(20_000, 10);
    const started = performance.now();
    await expect(loadWorkbook(fromBuffer(bytes), { contentLimits: { maxCells: 10 } })).rejects.toThrow(
      OpenXmlContentLimitError,
    );
    const elapsed = performance.now() - started;
    // Generous by two orders of magnitude against modelling all 200k cells;
    // this asserts the shape of the cost, not a machine's speed.
    expect(elapsed).toBeLessThan(2_000);
  });
});

describe('loadWorkbook with a row cap', () => {
  it('reads a workbook that fits and refuses one that does not', async () => {
    const bytes = await archive(10, 2);
    await expect(loadWorkbook(fromBuffer(bytes), { contentLimits: { maxRows: 10 } })).resolves.toBeDefined();
    await expect(loadWorkbook(fromBuffer(bytes), { contentLimits: { maxRows: 9 } })).rejects.toThrow(
      'worksheet: reading row 10 of Data passes contentLimits.maxRows of 9',
    );
  });
});

describe('a cap that cannot mean anything', () => {
  it('is rejected where the caller passes it', async () => {
    const bytes = await archive(2, 2);
    for (const bad of [0, -1, 1.5, Number.NaN, Number.POSITIVE_INFINITY]) {
      await expect(loadWorkbook(fromBuffer(bytes), { contentLimits: { maxCells: bad } })).rejects.toThrow(
        OpenXmlError,
      );
    }
    await expect(loadWorkbook(fromBuffer(bytes), { contentLimits: { maxCells: 0 } })).rejects.toThrow(
      'contentLimits.maxCells must be a positive integer; got 0',
    );
    await expect(loadWorkbook(fromBuffer(bytes), { contentLimits: { maxRows: -3 } })).rejects.toThrow(
      'contentLimits.maxRows must be a positive integer; got -3',
    );
  });
});

describe('loadWorkbookStream with a cell cap', () => {
  it('yields the rows that fit and then rejects the iterator', async () => {
    const wb = await loadWorkbookStream(fromBuffer(await archive(10, 2)), {
      contentLimits: { maxCells: 4 },
    });
    try {
      await expect(drain(wb.openWorksheet('Data'))).rejects.toThrow(
        'worksheet: reading Data!A3 passes contentLimits.maxCells of 4',
      );
    } finally {
      await wb.close();
    }
  });

  it('applies the cap on the indexed band-query path too', async () => {
    // minRow > 1 replays from the row-offset index instead of streaming the
    // part, which is a second route into the row iterator.
    const wb = await loadWorkbookStream(fromBuffer(await archive(10, 2)), {
      contentLimits: { maxCells: 4 },
    });
    try {
      await expect(drain(wb.openWorksheet('Data'), 5)).rejects.toThrow(OpenXmlContentLimitError);
    } finally {
      await wb.close();
    }
  });

  it('counts per traversal, so a second pass over the same sheet still runs', async () => {
    const wb = await loadWorkbookStream(fromBuffer(await archive(4, 2)), {
      contentLimits: { maxCells: 8 },
    });
    try {
      const ws = wb.openWorksheet('Data');
      expect(await drain(ws)).toHaveLength(4);
      expect(await drain(ws)).toHaveLength(4);
    } finally {
      await wb.close();
    }
  });

  it('reads everything when no cap is given', async () => {
    const wb = await loadWorkbookStream(fromBuffer(await archive(10, 2)));
    try {
      expect(await drain(wb.openWorksheet('Data'))).toHaveLength(10);
    } finally {
      await wb.close();
    }
  });
});

describe('loadWorkbookStream with a row cap', () => {
  it('rejects the iterator at the row past the cap', async () => {
    const wb = await loadWorkbookStream(fromBuffer(await archive(10, 2)), {
      contentLimits: { maxRows: 3 },
    });
    try {
      await expect(drain(wb.openWorksheet('Data'))).rejects.toThrow(
        'worksheet: reading row 4 of Data passes contentLimits.maxRows of 3',
      );
    } finally {
      await wb.close();
    }
  });
});

/** A one-sheet package whose `<sheetData>` is written verbatim. */
const rawPackage = (sheetData: string): Uint8Array => {
  const enc = new TextEncoder();
  const decl = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
  const relNs = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
  const pkgRelNs = 'http://schemas.openxmlformats.org/package/2006/relationships';
  const mainNs = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
  const ctNs = 'http://schemas.openxmlformats.org/package/2006/content-types';
  const sheetCt = 'application/vnd.openxmlformats-officedocument.spreadsheetml';
  return zipSync({
    '[Content_Types].xml': enc.encode(
      `${decl}<Types xmlns="${ctNs}">` +
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
        '<Default Extension="xml" ContentType="application/xml"/>' +
        `<Override PartName="/xl/workbook.xml" ContentType="${sheetCt}.sheet.main+xml"/>` +
        `<Override PartName="/xl/worksheets/sheet1.xml" ContentType="${sheetCt}.worksheet+xml"/></Types>`,
    ),
    '_rels/.rels': enc.encode(
      `${decl}<Relationships xmlns="${pkgRelNs}">` +
        `<Relationship Id="rId1" Type="${relNs}/officeDocument" Target="xl/workbook.xml"/></Relationships>`,
    ),
    'xl/workbook.xml': enc.encode(
      `${decl}<workbook xmlns="${mainNs}" xmlns:r="${relNs}">` +
        '<sheets><sheet name="Data" sheetId="1" r:id="rId1"/></sheets></workbook>',
    ),
    'xl/_rels/workbook.xml.rels': enc.encode(
      `${decl}<Relationships xmlns="${pkgRelNs}">` +
        `<Relationship Id="rId1" Type="${relNs}/worksheet" Target="worksheets/sheet1.xml"/></Relationships>`,
    ),
    'xl/worksheets/sheet1.xml': enc.encode(
      `${decl}<worksheet xmlns="${mainNs}"><sheetData>${sheetData}</sheetData></worksheet>`,
    ),
  });
};

describe('loadWorkbookStream with a cap that cannot mean anything', () => {
  it('rejects it from the call rather than from the iterator', async () => {
    // The budget used to be built inside the row generator, so this resolved
    // and the complaint arrived later as an iterator rejection, where a caller
    // cannot tell it from a workbook that was genuinely too big.
    const bytes = await archive(4, 2);
    await expect(loadWorkbookStream(fromBuffer(bytes), { contentLimits: { maxCells: 0 } })).rejects.toThrow(
      'contentLimits.maxCells must be a positive integer; got 0',
    );
    await expect(loadWorkbookStream(fromBuffer(bytes), { contentLimits: { maxRows: 1.5 } })).rejects.toThrow(
      OpenXmlError,
    );
  });

  it('rejects it even when the band asked for would yield nothing', async () => {
    // A band above every row returns an empty iterator that never builds a
    // budget, so validation cannot live in the generator.
    const bytes = await archive(4, 2);
    await expect(loadWorkbookStream(fromBuffer(bytes), { contentLimits: { maxCells: -1 } })).rejects.toThrow(
      'contentLimits.maxCells must be a positive integer; got -1',
    );
    const wb = await loadWorkbookStream(fromBuffer(bytes), { contentLimits: { maxCells: 8 } });
    try {
      expect(await drain(wb.openWorksheet('Data'), 99)).toHaveLength(0);
    } finally {
      await wb.close();
    }
  });
});

describe('loadWorkbookStream band queries', () => {
  it('charges the rows the index records, not only the ones the band yields', async () => {
    // Seeking to minRow means indexing every row of the part first, one object
    // per row, so a cap that only counted yielded rows left the largest
    // allocation on this path unbounded.
    const wb = await loadWorkbookStream(fromBuffer(await archive(10, 2)), {
      contentLimits: { maxRows: 5 },
    });
    try {
      await expect(drain(wb.openWorksheet('Data'), 9)).rejects.toThrow(
        'worksheet: reading row 6 of Data passes contentLimits.maxRows of 5',
      );
    } finally {
      await wb.close();
    }
  });

  it('yields the band when the part fits the cap', async () => {
    const wb = await loadWorkbookStream(fromBuffer(await archive(10, 2)), {
      contentLimits: { maxRows: 10 },
    });
    try {
      expect(await drain(wb.openWorksheet('Data'), 9)).toHaveLength(2);
    } finally {
      await wb.close();
    }
  });

  it('charges the rows it walks past on a sheet the index cannot seek', async () => {
    // No `@r` anywhere: the index cannot number these rows, so the band query
    // streams the part from the start and walks rows 1-3 to reach row 4.
    const rows = '<row><c><v>1</v></c></row>'.repeat(5);
    const wb = await loadWorkbookStream(fromBuffer(Buffer.from(rawPackage(rows))), {
      contentLimits: { maxRows: 2 },
    });
    try {
      await expect(drain(wb.openWorksheet('Data'), 4)).rejects.toThrow(
        'worksheet: reading an unnumbered row of Data passes contentLimits.maxRows of 2',
      );
    } finally {
      await wb.close();
    }
  });
});

describe('the shape of a band-query refusal', () => {
  it('comes from the iterRows call, since the index is built before a row is yielded', async () => {
    const wb = await loadWorkbookStream(fromBuffer(await archive(10, 2)), {
      contentLimits: { maxRows: 5 },
    });
    try {
      const ws = wb.openWorksheet('Data');
      expect(() => ws.iterRows({ minRow: 9 })).toThrow(OpenXmlContentLimitError);
    } finally {
      await wb.close();
    }
  });
});

describe('content limits before buffering unresolved coordinates', () => {
  for (const reader of ['eager', 'streaming'] as const) {
    const read = async (xml: string, limits: { maxCells?: number; maxRows?: number }): Promise<void> => {
      const bytes = rawPackage(xml);
      if (reader === 'eager') {
        await loadWorkbook(fromBuffer(bytes), { contentLimits: limits });
      } else {
        const wb = await loadWorkbookStream(fromBuffer(bytes), { contentLimits: limits });
        try {
          await drain(wb.openWorksheet('Data'));
        } finally {
          await wb.close();
        }
      }
    };

    it(`${reader}: stops at the cell cap before buffering a reference-less row`, async () => {
      // An invalid reference after the third cell proves that the limit stops
      // the walk before collecting the rest of this unresolved row.
      await expect(read('<row><c/><c/><c/><c r="invalid"/></row>', { maxCells: 2 }))
        .rejects.toThrow(OpenXmlContentLimitError);
    });

    it(`${reader}: stops at the row cap before reading an unnumbered row's cells`, async () => {
      await expect(read('<row/><row><c r="invalid"/></row>', { maxRows: 1 }))
        .rejects.toThrow(OpenXmlContentLimitError);
    });

    it(`${reader}: permits exactly the budget with derived coordinates`, async () => {
      await expect(read('<row><c/><c/></row><row><c/><c r="B5"/></row>', { maxCells: 4, maxRows: 2 }))
        .resolves.toBeUndefined();
    });
  }
});


describe('streaming cell budgets for filtered rows', () => {
  it('charges unresolved cells before their row can be excluded by the band', async () => {
    const wb = await loadWorkbookStream(fromBuffer(rawPackage('<row><c/><c/><c r="C9"/></row>')), {
      contentLimits: { maxCells: 1 },
    });
    try {
      const rows = async (): Promise<void> => {
        for await (const _row of wb.openWorksheet('Data').iterRows({ maxRow: 1 })) { /* drain */ }
      };
      await expect(rows()).rejects.toThrow(OpenXmlContentLimitError);
    } finally {
      await wb.close();
    }
  });

  it('does not charge cells in known excluded rows or columns', async () => {
    const wb = await loadWorkbookStream(fromBuffer(rawPackage(
      '<row r="1"><c r="A1"><v>1</v></c><c r="B1"><v>2</v></c></row>' +
      '<row r="2"><c r="A2"><v>3</v></c><c r="B2"><v>4</v></c></row>',
    )), { contentLimits: { maxCells: 1 } });
    try {
      const values: unknown[][] = [];
      for await (const row of wb.openWorksheet('Data').iterValues({ maxRow: 1, maxCol: 1 })) values.push(row);
      expect(values).toEqual([[1]]);
    } finally {
      await wb.close();
    }
  });
});
