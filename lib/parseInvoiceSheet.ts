import { GST_SLABS, isGstSlab } from "./calculateTax";
import {
  at,
  describeCell,
  isBlank,
  isBlankRow,
  toMultilineText,
  toNumber,
  toText,
} from "./cells";
import type {
  Cell,
  InvoiceMeta,
  InvoiceSheet,
  Issue,
  LineItem,
  ParseResult,
  SheetData,
} from "./types";

/** Column headings used in validation messages, in sheet order. */
const ITEM_COLUMNS = [
  "SI",
  "Description of Goods",
  "HSN/SAC",
  "Quantity",
  "Rate",
  "per",
  "GST %",
  "Amount",
] as const;

const SI = 0;
const DESCRIPTION = 1;
const HSN = 2;
const QUANTITY = 3;
const RATE = 4;
const PER = 5;
const GST = 6;
const AMOUNT = 7;

/** Data rows 0-5 hold the metadata and headings; line items start at data row 6. */
const FIRST_ITEM_ROW = 6;

/**
 * Read the invoice spreadsheet, which is addressed strictly by position:
 * alternating label/value rows, then a headings row, then one line item per row.
 * The spreadsheet's own heading row is dropped first, so data row 0 is sheet row 2.
 *
 * Parsing is deliberately tolerant — blank rows and trailing blank columns are
 * ignored — but every problem is reported against a sheet row and column rather
 * than surfacing as a stack trace (§7 bug 6).
 */
export function parseInvoiceSheet(rows: SheetData): ParseResult<InvoiceSheet> {
  const errors: Issue[] = [];
  const warnings: Issue[] = [];

  const body = trimTrailingBlankRows(rows.slice(1));

  if (body.length === 0) {
    errors.push({
      message:
        "The sheet is empty. Start from the sample sheet: row 1 holds the column labels and the values begin on row 2.",
    });
    return { data: null, errors, warnings };
  }

  if (body.length < FIRST_ITEM_ROW + 1) {
    errors.push({
      message: `The sheet has ${body.length + 1} rows but the invoice layout needs at least ${
        FIRST_ITEM_ROW + 2
      }: three label/value pairs, a headings row, then one row per line item.`,
    });
    return { data: null, errors, warnings };
  }

  const meta = readMeta(body, warnings);
  const items = readItems(body, errors, warnings);

  if (items.length === 0) {
    errors.push({
      message: `No line items found. Line items start on row ${FIRST_ITEM_ROW + 2}, below the headings row.`,
    });
  }

  if (errors.length > 0) return { data: null, errors, warnings };
  return { data: { meta, items }, errors, warnings };
}

function readMeta(body: SheetData, warnings: Issue[]): InvoiceMeta {
  const first = body[0];
  const second = body[2];
  const third = body[4];

  if (isBlank(at(first, 1))) {
    warnings.push({
      message: "Invoice number is blank.",
      location: describeCell(0, 1, "Invoice number"),
    });
  }

  const sameStateCell = at(third, 4);
  const sameStateText = toText(sameStateCell).trim().toLowerCase();
  if (sameStateText !== "yes" && sameStateText !== "no") {
    warnings.push({
      message: isBlank(sameStateCell)
        ? "Same-state flag is blank, so the invoice is treated as inter-state (IGST). Enter 'yes' for a CGST + SGST split."
        : `Same-state flag reads "${toText(sameStateCell)}", which is not 'yes' or 'no'. Anything other than 'yes' is treated as inter-state (IGST).`,
      location: describeCell(4, 4, "Same State"),
    });
  }

  return {
    invoiceType: toText(at(first, 0)) || "TAX INVOICE",
    invoiceNumber: toText(at(first, 1)),
    buyerName: toText(at(first, 2)),
    buyerAddress: toMultilineText(at(first, 3)),
    buyerPhone: toText(at(first, 4)),
    buyerGst: toText(at(first, 5)),
    billDate: toText(at(first, 6)),
    deliveryNote: toText(at(first, 7)),
    modeOfPayment: toText(at(second, 0)),
    supplierRef: toText(at(second, 1)),
    otherRef: toText(at(second, 2)),
    buyerOrderNumber: toText(at(second, 3)),
    buyerOrderDate: toText(at(second, 4)),
    despatchDocumentNumber: toText(at(second, 5)),
    deliveryNoteDate: toText(at(second, 6)),
    despatchedThrough: toText(at(second, 7)),
    destination: toMultilineText(at(third, 0)),
    term1: toText(at(third, 1)),
    term2: toText(at(third, 2)),
    term3: toText(at(third, 3)),
    sameState: sameStateText === "yes",
  };
}

function readItems(body: SheetData, errors: Issue[], warnings: Issue[]): LineItem[] {
  const items: LineItem[] = [];

  for (let rowIndex = FIRST_ITEM_ROW; rowIndex < body.length; rowIndex++) {
    const row = body[rowIndex];
    if (isBlankRow(row)) continue;

    const quantity = requireNumber(row, rowIndex, QUANTITY, errors, warnings);
    const amount = requireNumber(row, rowIndex, AMOUNT, errors, warnings);
    const rate = optionalNumber(row, rowIndex, RATE, errors);
    const gstPercent = readGstPercent(row, rowIndex, warnings);

    const description = toMultilineText(at(row, DESCRIPTION));
    if (description === "") {
      warnings.push({
        message: "Line item has no description.",
        location: describeCell(rowIndex, DESCRIPTION, ITEM_COLUMNS[DESCRIPTION]),
      });
    }

    items.push({
      si: toText(at(row, SI)),
      description,
      hsn: toText(at(row, HSN)),
      quantity: quantity ?? 0,
      rate,
      per: toText(at(row, PER)),
      gstPercent,
      amount: amount ?? 0,
    });
  }

  return items;
}

function requireNumber(
  row: Cell[],
  rowIndex: number,
  columnIndex: number,
  errors: Issue[],
  warnings: Issue[],
): number | null {
  const cell = at(row, columnIndex);
  const value = toNumber(cell);

  if (value === null) {
    warnings.push({
      message: `${ITEM_COLUMNS[columnIndex]} is blank and will be counted as 0.`,
      location: describeCell(rowIndex, columnIndex, ITEM_COLUMNS[columnIndex]),
    });
    return null;
  }

  if (Number.isNaN(value)) {
    errors.push({
      message: `${ITEM_COLUMNS[columnIndex]} must be a number, but reads "${toText(cell)}".`,
      location: describeCell(rowIndex, columnIndex, ITEM_COLUMNS[columnIndex]),
    });
    return null;
  }

  return value;
}

function optionalNumber(
  row: Cell[],
  rowIndex: number,
  columnIndex: number,
  errors: Issue[],
): number | null {
  const cell = at(row, columnIndex);
  const value = toNumber(cell);
  if (value === null) return null;

  if (Number.isNaN(value)) {
    errors.push({
      message: `${ITEM_COLUMNS[columnIndex]} must be a number, but reads "${toText(cell)}".`,
      location: describeCell(rowIndex, columnIndex, ITEM_COLUMNS[columnIndex]),
    });
    return null;
  }

  return value;
}

/**
 * A GST cell outside 5 / 12 / 18 is not fatal — the legacy app let the row through
 * untaxed — but it silently under-taxes the invoice, so it always raises a warning.
 */
function readGstPercent(row: Cell[], rowIndex: number, warnings: Issue[]): number | null {
  const cell = at(row, GST);
  const value = toNumber(cell);
  const location = describeCell(rowIndex, GST, ITEM_COLUMNS[GST]);
  const allowed = GST_SLABS.join(", ");

  if (value === null) {
    warnings.push({
      message: `GST % is blank, so this row is not taxed. Expected one of ${allowed}.`,
      location,
    });
    return null;
  }

  if (Number.isNaN(value)) {
    warnings.push({
      message: `GST % reads "${toText(cell)}", so this row is not taxed. Expected one of ${allowed}.`,
      location,
    });
    return null;
  }

  if (!isGstSlab(value)) {
    warnings.push({
      message: `GST % is ${value}, which is not a supported slab, so this row is not taxed. Expected one of ${allowed}.`,
      location,
    });
    return null;
  }

  return value;
}

export function trimTrailingBlankRows(rows: SheetData): SheetData {
  let end = rows.length;
  while (end > 0 && isBlankRow(rows[end - 1])) end--;
  return rows.slice(0, end);
}
