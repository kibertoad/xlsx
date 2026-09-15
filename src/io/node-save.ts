// Node-only counterpart to `workbookToBytes` in `./save`. Returns a Buffer
// directly so Node consumers don't pay a `Buffer.from(uint8Array)` copy.

import type { Workbook } from '../workbook/workbook.js';
import { toBuffer } from './node.js';
import { saveWorkbook, type SaveOptions } from './save.js';

export async function workbookToBuffer(wb: Workbook, opts?: SaveOptions): Promise<Buffer> {
  const sink = toBuffer();
  await saveWorkbook(wb, sink, opts);
  return sink.result();
}
