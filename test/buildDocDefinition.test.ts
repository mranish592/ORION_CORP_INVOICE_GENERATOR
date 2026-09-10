import assert from "node:assert/strict";
import { test } from "node:test";

import { COMPANY } from "../config/company";
import { buildInvoiceDocDefinition } from "../lib/buildInvoiceDocDefinition";
import {
  buildPackingDocDefinition,
  PACKING_PRODUCT_HEADER_ROW,
} from "../lib/buildPackingDocDefinition";
import { calculateInvoiceTotals } from "../lib/calculateTax";
import { parseInvoiceSheet } from "../lib/parseInvoiceSheet";
import { parsePackingSheet } from "../lib/parsePackingSheet";
import { CONTENT_WIDTH, drawnTableWidth } from "../lib/pdfLayouts";
import { readSample } from "./helpers";

/** Document definitions are deeply polymorphic; these tests poke at them structurally. */
// eslint-disable-next-line
type Node = any;

const invoiceSheet = parseInvoiceSheet(await readSample("invoice_sample.xlsx")).data!;
const packingSheet = parsePackingSheet(await readSample("packing_sample.xlsx")).data!;

function invoiceDoc(sameState: boolean) {
  const sheet = { ...invoiceSheet, meta: { ...invoiceSheet.meta, sameState } };
  return buildInvoiceDocDefinition({
    sheet,
    totals: calculateInvoiceTotals(sheet.items, sameState),
    signature: "data:image/jpeg;base64,AAAA",
  });
}

/** Every string in the document, flattened, for whole-document assertions. */
function allText(node: Node): string[] {
  if (node === null || node === undefined) return [];
  if (typeof node === "string") return [node];
  if (Array.isArray(node)) return node.flatMap(allText);
  if (typeof node === "object") {
    return Object.entries(node)
      .filter(([key]) => key !== "layout")
      .flatMap(([, value]) => allText(value));
  }
  return [];
}

function tables(doc: Node): Node[] {
  return (doc.content as Node[]).filter((node) => node && node.table);
}

test("every invoice table draws exactly the 540pt content width", () => {
  // pdfmake adds cell padding and rule widths outside the declared widths, so the
  // drawn width — not the sum of `widths` — is what has to fit the page.
  for (const node of tables(invoiceDoc(false))) {
    assert.ok(Math.abs(drawnTableWidth(node.table.widths as number[]) - CONTENT_WIDTH) < 1e-9);
  }
});

test("stacks the seven blocks in the order the legacy layout used", () => {
  const doc = invoiceDoc(false);
  const content = doc.content as Node[];
  assert.equal(content.length, 8);
  assert.equal(content[0].text, invoiceSheet.meta.invoiceType);
  assert.equal(content[7].text, COMPANY.footer);
  assert.deepEqual(
    tables(doc).map((node) => (node.table.widths as number[]).length),
    [3, 2, 8, 2, 5, 2],
  );
});

test("inter-state invoices use the five-column IGST tax summary", () => {
  const doc = invoiceDoc(false);
  const taxTable = tables(doc)[4];

  assert.equal(taxTable.table.widths.length, 5);
  const text = allText(taxTable);
  assert.ok(text.includes("Integrated Tax"));
  assert.equal(text.includes("Central Tax"), false);
  assert.equal(text.includes("State Tax"), false);
});

test("intra-state invoices use the seven-column CGST/SGST tax summary", () => {
  const doc = invoiceDoc(true);
  const taxTable = tables(doc)[4];

  assert.equal(taxTable.table.widths.length, 7);
  const text = allText(taxTable);
  assert.ok(text.includes("Central Tax"));
  assert.ok(text.includes("State Tax"));
  assert.equal(text.includes("Integrated Tax"), false);
});

test("inter-state line items carry one IGST row per slab", () => {
  const rows = allText(tables(invoiceDoc(false))[2]);
  assert.ok(rows.includes("IGST 5%"));
  assert.ok(rows.includes("IGST 18%"));
  assert.equal(rows.includes("CGST 2.5%"), false);
});

test("intra-state line items split each slab into CGST and SGST rows", () => {
  const rows = allText(tables(invoiceDoc(true))[2]);
  assert.ok(rows.includes("CGST 2.5%"));
  assert.ok(rows.includes("SGST 2.5%"));
  assert.ok(rows.includes("CGST 9%"));
  assert.ok(rows.includes("SGST 9%"));
  assert.equal(rows.includes("IGST 5%"), false);
});

test("prints the HSN codes as a joined list, not a raw set", () => {
  // Legacy bug §7.1: Python sets rendered as "{3921, 6307, 40151900}".
  const text = allText(tables(invoiceDoc(false))[4]);
  assert.ok(text.includes("3921, 40151900, 6307"));
  assert.equal(
    text.some((value) => value.includes("{")),
    false,
  );
});

test("the tax amount in words ends in a single 'Only'", () => {
  // Legacy bug §7.2: the declaration appended ' only' to a string already
  // ending in ' Paisa Only'.
  const line = allText(tables(invoiceDoc(false))[5]).find((value) =>
    value.startsWith("Tax Amount (in words):"),
  );
  assert.ok(line);
  assert.ok(line.endsWith("Paisa Only"));
  assert.equal(/only/i.test(line.replace(/Paisa Only$/, "")), false);
});

test("prints the corrected 'for Orion Corp' signature line", () => {
  const text = allText(tables(invoiceDoc(false))[5]);
  assert.ok(text.includes("for Orion Corp"));
  assert.equal(text.includes("for Orion the Corp"), false);
});

test("embeds the signature only when one is supplied", () => {
  const withSignature = invoiceDoc(false);
  assert.deepEqual(Object.keys(withSignature.images ?? {}), ["signature"]);

  const withoutSignature = buildInvoiceDocDefinition({
    sheet: invoiceSheet,
    totals: calculateInvoiceTotals(invoiceSheet.items, false),
    signature: null,
  });
  assert.equal(withoutSignature.images, undefined);
  assert.equal(
    JSON.stringify(withoutSignature.content).includes('"image":"signature"'),
    false,
  );
});

test("uses the rupee symbol in figures and INR in the amount in words", () => {
  const text = allText(invoiceDoc(false));
  assert.ok(text.includes("₹3,55,605.00"));
  assert.ok(
    text.some((value) => value.startsWith("INR ") && value.endsWith(" Paisa Only")),
  );
});

test("repeats the line-item header across pages", () => {
  assert.equal(tables(invoiceDoc(false))[2].table.headerRows, 1);
});

test("builds one unbreakable card per packing box, sized to the content width", () => {
  const doc = buildPackingDocDefinition(packingSheet);
  const cards = doc.content as Node[];

  assert.equal(cards.length, 5);
  for (const card of cards) {
    assert.equal(card.unbreakable, true);
    assert.ok(Math.abs(drawnTableWidth(card.table.widths as number[]) - CONTENT_WIDTH) < 1e-9);
  }
});

test("a packing card lists only the products the box actually contains", () => {
  const doc = buildPackingDocDefinition(packingSheet);
  const rows = (doc.content as Node[])[1].table.body as Node[][];
  const productRows = rows.slice(PACKING_PRODUCT_HEADER_ROW);

  assert.deepEqual(productRows[0].map((cell: Node) => cell.text), ["Product", "Quantity"]);
  assert.equal(productRows.length, 2);
  assert.deepEqual(productRows[1], ["Blue Respirator", "100"]);
});

test("draws rules only from the product header row down", () => {
  const doc = buildPackingDocDefinition(packingSheet);
  const rows = (doc.content as Node[])[0].table.body as Node[][];

  for (const cell of rows.slice(0, PACKING_PRODUCT_HEADER_ROW).flat()) {
    if (Object.keys(cell).length === 0) continue; // colSpan placeholder
    assert.deepEqual(cell.border, [false, false, false, false]);
  }
  for (const cell of rows.slice(PACKING_PRODUCT_HEADER_ROW).flat()) {
    assert.equal(cell.border, undefined);
  }
});

test("repeats the invoice metadata on every packing card", () => {
  const doc = buildPackingDocDefinition(packingSheet);
  for (const card of doc.content as Node[]) {
    const text = allText(card.table.body.slice(0, PACKING_PRODUCT_HEADER_ROW));
    assert.ok(text.includes("Invoice no: 72"));
    assert.ok(text.includes("Date: 31/07/2020"));
    assert.ok(text.includes("Contact No: 9003840110"));
  }
});
