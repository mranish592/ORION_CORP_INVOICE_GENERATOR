import type { Cell } from "./types";

/** Spreadsheet column letter for a zero-based column index: 0 -> "A", 26 -> "AA". */
export function columnLetter(index: number): string {
  let n = index;
  let letters = "";
  do {
    letters = String.fromCharCode(65 + (n % 26)) + letters;
    n = Math.floor(n / 26) - 1;
  } while (n >= 0);
  return letters;
}

/**
 * Both parsers drop the spreadsheet's own heading row before indexing, mirroring
 * the way pandas consumed it in the legacy app. Data row 0 is therefore sheet row 2.
 */
export function sheetRowNumber(dataRowIndex: number): number {
  return dataRowIndex + 2;
}

/** "row 8, column G (GST %)" — the location shown next to a validation message. */
export function describeCell(
  dataRowIndex: number,
  columnIndex: number,
  columnName?: string,
): string {
  const ref = `row ${sheetRowNumber(dataRowIndex)}, column ${columnLetter(columnIndex)}`;
  return columnName ? `${ref} (${columnName})` : ref;
}

export function isBlank(cell: Cell): boolean {
  if (cell === null || cell === undefined) return true;
  return typeof cell === "string" && cell.trim() === "";
}

/** Dates arrive from `read-excel-file` at UTC midnight, so read them back in UTC. */
export function formatSheetDate(date: Date): string {
  const day = String(date.getUTCDate()).padStart(2, "0");
  const month = String(date.getUTCMonth() + 1).padStart(2, "0");
  return `${day}/${month}/${date.getUTCFullYear()}`;
}

/**
 * Coerce any cell to display text. Blank cells become `''`, which is what the
 * legacy pandas pipeline did via `replace(np.nan, '')`; doing it here is what
 * stops `'Buyer Name: ' + cell` from throwing on a numeric cell (§7 bug 5).
 */
export function toText(cell: Cell): string {
  if (cell === null || cell === undefined) return "";
  if (cell instanceof Date) return formatSheetDate(cell);
  if (typeof cell === "number") return Number.isFinite(cell) ? String(cell) : "";
  if (typeof cell === "boolean") return cell ? "yes" : "no";
  return cell;
}

/**
 * Same as {@link toText}, but turns the literal two-character sequence `\n` into a
 * real line break. Addresses in the sample sheets were typed that way and the
 * legacy PDF swallowed them; honouring them keeps multi-line addresses readable.
 */
export function toMultilineText(cell: Cell): string {
  return toText(cell).replace(/\\r\\n|\\n/g, "\n");
}

/** `null` for a blank cell, `NaN` for a cell that is not a number. */
export function toNumber(cell: Cell): number | null {
  if (isBlank(cell)) return null;
  if (typeof cell === "number") return cell;
  if (cell instanceof Date) return NaN;
  if (typeof cell === "boolean") return NaN;
  const parsed = Number(String(cell).replace(/,/g, "").trim());
  return Number.isFinite(parsed) ? parsed : NaN;
}

/** Read a cell that may be missing entirely because the row is short. */
export function at(row: Cell[] | undefined, index: number): Cell {
  return row ? row[index] : null;
}

export function isBlankRow(row: Cell[] | undefined): boolean {
  if (!row) return true;
  return row.every(isBlank);
}

/** Round to 2 decimals without the float drift of `Math.round(n * 100) / 100`. */
export function round2(value: number): number {
  return Number((Math.round((value + Number.EPSILON) * 100) / 100).toFixed(2));
}
