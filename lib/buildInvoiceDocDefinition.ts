import type { Content, TableCell, TDocumentDefinitions } from "pdfmake/interfaces";

import { COMPANY } from "../config/company";
import { amountInWords } from "./amountInWords";
import { round2 } from "./cells";
import { taxRowsFor } from "./calculateTax";
import { formatMoney, formatQuantity } from "./format";
import {
  boxLayout,
  buyerTermsLayout,
  contentWidths,
  declarationLayout,
  gridLayout,
  lineItemsLayout,
  PAGE_SIDE_MARGIN,
} from "./pdfLayouts";
import type { InvoiceMeta, InvoiceSheet, InvoiceTotals } from "./types";

export interface InvoiceDocOptions {
  sheet: InvoiceSheet;
  totals: InvoiceTotals;
  /** `data:image/jpeg;base64,...` for the authorised-signature scan. */
  signature?: string | null;
}

const NO_MARGIN: [number, number, number, number] = [0, 0, 0, 0];

/**
 * Build the pdfmake document definition for a GST tax invoice.
 *
 * Pure: no React, no pdfmake runtime, no I/O — which is what makes the layout
 * testable and keeps the totals out of shared mutable state (§7 bug 3).
 */
export function buildInvoiceDocDefinition(options: InvoiceDocOptions): TDocumentDefinitions {
  const { sheet, totals, signature } = options;
  const { meta } = sheet;

  const content: Content[] = [
    {
      text: meta.invoiceType,
      alignment: "center",
      fontSize: 12,
      margin: [0, 0, 0, 20],
    },
    fixedDetailsTable(meta),
    buyerAndTermsTable(meta),
    lineItemsTable(sheet, totals),
    amountChargeableTable(totals),
    taxSummaryTable(totals),
    declarationTable(totals, signature),
    {
      text: COMPANY.footer,
      alignment: "center",
      fontSize: 10,
      margin: [0, 6, 0, 0],
    },
  ];

  return {
    pageSize: "A4",
    pageMargins: [PAGE_SIDE_MARGIN, 20, PAGE_SIDE_MARGIN, 20],
    defaultStyle: { fontSize: 8, lineHeight: 1.15 },
    info: {
      title: `${meta.invoiceType} ${meta.invoiceNumber}`.trim(),
      author: COMPANY.name,
    },
    images: signature ? { signature } : undefined,
    content,
  };
}

/** `label\nvalue`, the paired cell shape used all through the header block. */
function labelled(label: string, value: string): string {
  return value === "" ? label : `${label}\n${value}`;
}

/** §6.1 — company block spanning all six rows, paired label/value cells beside it. */
function fixedDetailsTable(meta: InvoiceMeta): Content {
  const companyBlock: Content = {
    fontSize: 8,
    text: [
      { text: COMPANY.name, bold: true },
      `\n${COMPANY.address}`,
      `\n${COMPANY.warehouse}`,
      `\nGSTIN/UIN: ${COMPANY.gstin}`,
      `\n${COMPANY.state}\n`,
      { text: "\nBank Details", bold: true },
      `\n${COMPANY.bank.accountName}`,
      `\n${COMPANY.bank.line}`,
      `\n${COMPANY.bank.ifsc}`,
      "\nBranch Address",
      `\n${COMPANY.bank.branchAddress}`,
    ],
  };

  const body: TableCell[][] = [
    [
      { ...(companyBlock as object), rowSpan: 6 } as TableCell,
      labelled("Inv No.", meta.invoiceNumber),
      labelled("Dated", meta.billDate),
    ],
    [{}, labelled("Delivery Note", meta.deliveryNote), labelled("Mode/Terms of Payment", meta.modeOfPayment)],
    [{}, labelled("Other Reference(s)", meta.otherRef), labelled("Buyer's Order No.", meta.buyerOrderNumber)],
    [{}, labelled("Order Date:", meta.buyerOrderDate), labelled("Despatch Document No.", meta.despatchDocumentNumber)],
    [{}, labelled("Delivery Note Date", meta.deliveryNoteDate), labelled("Despatch through", meta.despatchedThrough)],
    [{}, { text: `SHIP TO: ${meta.destination}`, colSpan: 2 }, {}],
  ];

  return {
    margin: NO_MARGIN,
    table: { widths: contentWidths([280, 130, 130]), body },
    layout: gridLayout,
  };
}

/** §6.2 — buyer identity on the left, delivery terms on the right. */
function buyerAndTermsTable(meta: InvoiceMeta): Content {
  const term = (value: string) => (value === "" ? "" : `*${value}`);

  const body: TableCell[][] = [
    [`Buyer Name: ${meta.buyerName}`, "Terms of Delivery"],
    [{ text: `Address: ${meta.buyerAddress}`, alignment: "justify" }, term(meta.term1)],
    [`Phone no.: ${meta.buyerPhone}`, term(meta.term2)],
    [`GST/PAN no.: ${meta.buyerGst}`, term(meta.term3)],
  ];

  return {
    margin: NO_MARGIN,
    table: { widths: contentWidths([280, 260]), body },
    layout: buyerTermsLayout,
  };
}

/** §6.3 — line items, then the subtotal, the per-slab tax rows and the total. */
function lineItemsTable(sheet: InvoiceSheet, totals: InvoiceTotals): Content {
  const header: TableCell[] = [
    "SI",
    "Description of Goods",
    "HSN/SAC",
    "Quantity",
    "Rate",
    "per",
    "GST%",
    "Amount",
  ].map((text) => ({ text, bold: true, alignment: "center" as const }));

  const body: TableCell[][] = [header];

  for (const item of sheet.items) {
    body.push([
      { text: item.si, alignment: "center" },
      { text: item.description, fontSize: 6, alignment: "center" },
      { text: item.hsn, alignment: "center" },
      { text: formatQuantity(item.quantity), alignment: "center" },
      { text: item.rate === null ? "" : formatMoney(item.rate), alignment: "right" },
      { text: item.per, alignment: "center" },
      { text: item.gstPercent === null ? "" : String(item.gstPercent), alignment: "center" },
      { text: formatMoney(item.amount), alignment: "right" },
    ]);
  }

  const summaryStartRow = body.length;

  const summaryRow = (
    label: string,
    value: string,
    quantity = "",
  ): TableCell[] => [
    "",
    { text: label, alignment: "right" },
    "",
    { text: quantity, alignment: "center" },
    "",
    { text: label === "" ? "" : "%", alignment: "center" },
    "",
    { text: value, alignment: "right" },
  ];

  body.push(summaryRow("", formatMoney(totals.taxableAmount)));

  for (const row of taxRowsFor(totals)) {
    body.push(summaryRow(row.label, formatMoney(row.amount)));
  }

  const totalRow = summaryRow("Total", formatMoney(totals.totalAmount), formatQuantity(totals.totalPieces));
  totalRow[1] = { text: "Total", alignment: "right", bold: true };
  totalRow[7] = { text: formatMoney(totals.totalAmount), alignment: "right", bold: true };
  body.push(totalRow);

  return {
    margin: NO_MARGIN,
    table: {
      widths: contentWidths([20, 210, 60, 60, 50, 30, 40, 70]),
      headerRows: 1,
      body,
    },
    layout: lineItemsLayout(summaryStartRow),
  };
}

/** §6.4 — the amount in words, spanning the full width. */
function amountChargeableTable(totals: InvoiceTotals): Content {
  const body: TableCell[][] = [
    ["Amount chargeable (inwords)", { text: "E. &O.E", alignment: "right" }],
    [{ text: amountInWords(totals.totalAmount), colSpan: 2, bold: true }, {}],
  ];

  return {
    margin: NO_MARGIN,
    table: { widths: contentWidths([470, 70]), body },
    layout: boxLayout,
  };
}

/** §6.5 — seven columns for an intra-state CGST/SGST split, five for IGST. */
function taxSummaryTable(totals: InvoiceTotals): Content {
  const centred = (text: string): TableCell => ({ text, alignment: "center", bold: true });

  if (totals.sameState) {
    const body: TableCell[][] = [
      [
        { text: "HSN/SAC", rowSpan: 2, bold: true },
        { text: "Taxable Value", rowSpan: 2, bold: true, alignment: "center" },
        { text: "Central Tax", colSpan: 2, bold: true, alignment: "center" },
        {},
        { text: "State Tax", colSpan: 2, bold: true, alignment: "center" },
        {},
        { text: "Total Tax Amount", rowSpan: 2, bold: true, alignment: "center" },
      ],
      [{}, {}, centred("Rate"), centred("Amount"), centred("Rate"), centred("Amount"), {}],
    ];

    for (const slab of totals.slabs) {
      const half = round2(slab.tax / 2);
      const halfRate = `${slab.rate / 2}%`;
      body.push([
        slab.hsnCodes.join(", "),
        { text: formatMoney(slab.taxableValue), alignment: "right" },
        { text: halfRate, alignment: "center" },
        { text: formatMoney(half), alignment: "right" },
        { text: halfRate, alignment: "center" },
        { text: formatMoney(half), alignment: "right" },
        { text: formatMoney(slab.tax), alignment: "right" },
      ]);
    }

    const halfTotal = round2(totals.totalTax / 2);
    body.push([
      { text: "Total", bold: true },
      { text: formatMoney(totals.taxableAmount), alignment: "right", bold: true },
      "",
      { text: formatMoney(halfTotal), alignment: "right", bold: true },
      "",
      { text: formatMoney(halfTotal), alignment: "right", bold: true },
      { text: formatMoney(totals.totalTax), alignment: "right", bold: true },
    ]);

    return {
      margin: NO_MARGIN,
      table: { widths: contentWidths([120, 80, 40, 80, 40, 80, 100]), headerRows: 2, body },
      layout: gridLayout,
    };
  }

  const body: TableCell[][] = [
    [
      { text: "HSN/SAC", rowSpan: 2, bold: true },
      { text: "Taxable Value", rowSpan: 2, bold: true, alignment: "center" },
      { text: "Integrated Tax", colSpan: 2, bold: true, alignment: "center" },
      {},
      { text: "Total Tax Amount", rowSpan: 2, bold: true, alignment: "center" },
    ],
    [{}, {}, centred("Rate"), centred("Amount"), {}],
  ];

  for (const slab of totals.slabs) {
    body.push([
      slab.hsnCodes.join(", "),
      { text: formatMoney(slab.taxableValue), alignment: "right" },
      { text: `${slab.rate}%`, alignment: "center" },
      { text: formatMoney(slab.tax), alignment: "right" },
      { text: formatMoney(slab.tax), alignment: "right" },
    ]);
  }

  body.push([
    { text: "Total", bold: true },
    { text: formatMoney(totals.taxableAmount), alignment: "right", bold: true },
    "",
    { text: formatMoney(totals.totalTax), alignment: "right", bold: true },
    { text: formatMoney(totals.totalTax), alignment: "right", bold: true },
  ]);

  return {
    margin: NO_MARGIN,
    table: { widths: contentWidths([240, 80, 40, 80, 100]), headerRows: 2, body },
    layout: gridLayout,
  };
}

/**
 * §6.6 — tax amount in words, declaration and signature.
 *
 * `amountInWords` already ends in "Paisa Only", so nothing appends a second
 * "only" the way the legacy code did (§7 bug 2).
 */
function declarationTable(totals: InvoiceTotals, signature?: string | null): Content {
  const signatureCell: TableCell = signature
    ? { image: "signature", width: 100, height: 60, alignment: "right" }
    : { text: "", alignment: "right" };

  const body: TableCell[][] = [
    [{ text: `Tax Amount (in words): ${amountInWords(totals.totalTax)}`, colSpan: 2 }, {}],
    [{ text: "Declaration", bold: true }, { text: COMPANY.signatureFor, alignment: "right" }],
    [
      `MSME Number: ${COMPANY.msmeNumber}\nGeM Seller ID: ${COMPANY.gemSellerId}`,
      signatureCell,
    ],
    [COMPANY.declaration, { text: "Authorised Signatory", alignment: "right" }],
  ];

  return {
    margin: NO_MARGIN,
    table: {
      widths: contentWidths([280, 260]),
      heights: [20, "auto", "auto", "auto"],
      body,
    },
    layout: declarationLayout,
  };
}
