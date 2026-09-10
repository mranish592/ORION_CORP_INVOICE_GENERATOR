import { at, describeCell, isBlankRow, toMultilineText, toNumber, toText } from "./cells";
import { trimTrailingBlankRows } from "./parseInvoiceSheet";
import type { Issue, PackingBox, PackingLine, PackingSheet, ParseResult, SheetData } from "./types";

/** Data row 5 names the products; boxes start at data row 6. */
const PRODUCT_HEADING_ROW = 5;
const FIRST_BOX_ROW = 6;

/**
 * Read the packing spreadsheet. Like the invoice sheet it is addressed by
 * position: three label/value pairs, a product-name row, then one row per box.
 *
 * A quantity of 0 (or a blank cell, which the legacy pandas pipeline coerced to 0)
 * means the product is simply omitted from that box's card.
 */
export function parsePackingSheet(rows: SheetData): ParseResult<PackingSheet> {
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

  if (body.length < FIRST_BOX_ROW + 1) {
    errors.push({
      message: `The sheet has ${body.length + 1} rows but the packing layout needs at least ${
        FIRST_BOX_ROW + 2
      }: three label/value pairs, a product-name row, then one row per box.`,
    });
    return { data: null, errors, warnings };
  }

  const products = readProductNames(body[PRODUCT_HEADING_ROW], warnings);
  if (products.length === 0) {
    errors.push({
      message: "No product names found. Name each product across the product row, starting in column B.",
      location: describeCell(PRODUCT_HEADING_ROW, 1, "Product names"),
    });
  }

  const boxes = readBoxes(body, products, errors, warnings);
  if (boxes.length === 0 && errors.length === 0) {
    errors.push({
      message: `No boxes found. Box rows start on row ${FIRST_BOX_ROW + 2}, below the product row.`,
    });
  }

  if (errors.length > 0) return { data: null, errors, warnings };

  return {
    data: {
      invoiceNumber: toText(at(body[0], 0)),
      shippingAddress: toMultilineText(at(body[0], 1)),
      buyerName: toText(at(body[2], 0)),
      date: toText(at(body[2], 1)),
      contactNumber: toText(at(body[4], 0)),
      boxes,
    },
    errors,
    warnings,
  };
}

interface ProductColumn {
  columnIndex: number;
  name: string;
}

/** Column 0 of the product row is the "Box" label; the rest are product names. */
function readProductNames(row: SheetData[number], warnings: Issue[]): ProductColumn[] {
  const products: ProductColumn[] = [];
  if (!row) return products;

  for (let columnIndex = 1; columnIndex < row.length; columnIndex++) {
    const name = toText(row[columnIndex]).trim();
    if (name === "") continue;
    products.push({ columnIndex, name });
  }

  if (products.length === 0) {
    warnings.push({
      message: "The product row names no products.",
      location: describeCell(PRODUCT_HEADING_ROW, 1, "Product names"),
    });
  }

  return products;
}

function readBoxes(
  body: SheetData,
  products: ProductColumn[],
  errors: Issue[],
  warnings: Issue[],
): PackingBox[] {
  const boxes: PackingBox[] = [];

  for (let rowIndex = FIRST_BOX_ROW; rowIndex < body.length; rowIndex++) {
    const row = body[rowIndex];
    if (isBlankRow(row)) continue;

    const lines: PackingLine[] = [];
    for (const product of products) {
      const cell = at(row, product.columnIndex);
      const quantity = toNumber(cell);

      if (quantity === null) continue; // blank means "not in this box"

      if (Number.isNaN(quantity)) {
        errors.push({
          message: `Quantity for "${product.name}" must be a number, but reads "${toText(cell)}".`,
          location: describeCell(rowIndex, product.columnIndex, product.name),
        });
        continue;
      }

      if (quantity === 0) continue; // 0 means "not in this box"
      lines.push({ product: product.name, quantity });
    }

    const boxNumber = toText(at(row, 0));
    if (lines.length === 0) {
      warnings.push({
        message: `Box ${boxNumber || rowIndex - FIRST_BOX_ROW + 1} has no products with a quantity above 0; its card will be empty.`,
        location: describeCell(rowIndex, 0, "Box"),
      });
    }

    boxes.push({ boxNumber, lines });
  }

  return boxes;
}
