import { readSheet } from "read-excel-file/node";

import type { SheetData } from "../lib/types";

/** Read one of the sample workbooks the way the browser does, but from disk. */
export async function readSample(name: string): Promise<SheetData> {
  const rows = await readSheet(`public/samples/${name}`);
  return rows as SheetData;
}
