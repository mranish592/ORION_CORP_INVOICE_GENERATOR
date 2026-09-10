import type { SheetData } from "./types";

/**
 * Read the first sheet of an uploaded workbook, entirely in the browser.
 *
 * `read-excel-file` is used rather than SheetJS `xlsx`, which is frozen on npm at
 * 0.18.5 with two unfixed high-severity advisories and no upgrade path.
 */
export async function readSpreadsheet(file: File): Promise<SheetData> {
  const { readSheet } = await import("read-excel-file/browser");
  const rows = await readSheet(file);
  return rows as SheetData;
}
