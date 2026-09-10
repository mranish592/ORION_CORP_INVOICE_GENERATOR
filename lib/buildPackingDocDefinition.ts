import type { Content, TableCell, TDocumentDefinitions } from "pdfmake/interfaces";

import { COMPANY } from "../config/company";
import { formatQuantity } from "./format";
import { contentWidths, gridLayout, PAGE_SIDE_MARGIN } from "./pdfLayouts";
import type { PackingBox, PackingSheet } from "./types";

const CARD_WIDTHS = contentWidths([300, 240]);

/**
 * The card's header block carries no rules — only the product grid below it is
 * bordered. Suppressing the borders per cell keeps the whole card in one table, so
 * the two halves share a column grid exactly.
 */
const NO_BORDER: [boolean, boolean, boolean, boolean] = [false, false, false, false];

/** Index of the `Product | Quantity` row, where the borders start. */
const PRODUCT_HEADER_ROW = 5;

/**
 * Build the pdfmake document definition for a packing list: one card per box.
 *
 * The legacy code paginated with a hand-rolled space budget that guessed at row
 * heights — 42 units per page, each card costing 10 plus the address length over
 * 50 plus its product count. pdfmake measures for real, so each card is simply
 * marked `unbreakable` and the natural flow decides where the page ends.
 */
export function buildPackingDocDefinition(sheet: PackingSheet): TDocumentDefinitions {
  return {
    pageSize: "A4",
    pageMargins: [PAGE_SIDE_MARGIN, 40, PAGE_SIDE_MARGIN, 40],
    defaultStyle: { fontSize: 10, lineHeight: 1.15 },
    info: {
      title: `Packing list ${sheet.invoiceNumber}`.trim(),
      author: COMPANY.name,
    },
    content: sheet.boxes.map((box) => boxCard(box, sheet)),
  };
}

function boxCard(box: PackingBox, sheet: PackingSheet): Content {
  const body: TableCell[][] = [
    [{ text: `Box No: ${box.boxNumber}`, colSpan: 2, bold: true, border: NO_BORDER }, {}],
    [
      { text: `Invoice no: ${sheet.invoiceNumber}`, border: NO_BORDER },
      { text: `Date: ${sheet.date}`, alignment: "right", border: NO_BORDER },
    ],
    [
      { text: `Buyer: ${sheet.buyerName}`, border: NO_BORDER },
      {
        text: `Contact No: ${sheet.contactNumber}`,
        alignment: "right",
        border: NO_BORDER,
      },
    ],
    [{ text: `Address: ${sheet.shippingAddress}`, colSpan: 2, border: NO_BORDER }, {}],
    [{ text: " ", colSpan: 2, border: NO_BORDER }, {}],
    [
      { text: "Product", bold: true },
      { text: "Quantity", bold: true },
    ],
    ...box.lines.map((line): TableCell[] => [line.product, formatQuantity(line.quantity)]),
  ];

  return {
    unbreakable: true,
    margin: [0, 0, 0, 40],
    table: { widths: CARD_WIDTHS, body },
    layout: gridLayout,
  };
}

/** Exported so tests can assert where the bordered product grid begins. */
export const PACKING_PRODUCT_HEADER_ROW = PRODUCT_HEADER_ROW;
