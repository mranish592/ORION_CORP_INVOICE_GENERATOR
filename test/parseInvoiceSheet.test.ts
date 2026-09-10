import assert from "node:assert/strict";
import { test } from "node:test";

import { parseInvoiceSheet } from "../lib/parseInvoiceSheet";
import type { SheetData } from "../lib/types";
import { readSample } from "./helpers";

const sample = await readSample("invoice_sample.xlsx");

test("reads the metadata rows by position", () => {
  const { data, errors } = parseInvoiceSheet(sample);
  assert.deepEqual(errors, []);
  assert.ok(data);

  const { meta } = data;
  assert.equal(meta.invoiceType, "PROFORMA INVOICE");
  assert.equal(meta.invoiceNumber, "23");
  assert.equal(meta.buyerName, "OBAT MEDICARE");
  assert.equal(meta.buyerGst, "10AAECB2157Q1ZP");
  assert.equal(meta.billDate, "21/01/20 12:00 AM");
  assert.equal(meta.otherRef, "2");
  assert.equal(meta.term1, "Delivery schedule will be alloted after receipt of payment");
  assert.equal(meta.sameState, false);
});

test("coerces numeric cells to text rather than throwing on concatenation", () => {
  const { data } = parseInvoiceSheet(sample);
  // Legacy bug §7.5: `'Phone no.: ' + cell` crashed when the cell held a number.
  assert.equal(data!.meta.buyerPhone, "8210064245");
});

test("formats date-typed cells as dd/mm/yyyy in UTC", () => {
  const { data } = parseInvoiceSheet(sample);
  assert.equal(data!.meta.buyerOrderDate, "05/05/2020");
  assert.equal(data!.meta.deliveryNoteDate, "05/06/2020");
});

test("reads every line item and keeps HSN codes as integers", () => {
  const { data } = parseInvoiceSheet(sample);
  const items = data!.items;

  assert.equal(items.length, 7);
  assert.deepEqual(
    items.map((item) => item.hsn),
    ["6307", "6307", "6307", "6307", "3921", "40151900", "6307"],
  );
  assert.deepEqual(
    items.map((item) => item.gstPercent),
    [5, 5, 5, 5, 18, 18, 18],
  );
  assert.deepEqual(
    items.map((item) => item.amount),
    [200000, 70000, 4200, 2100, 21000, 4500, 30000],
  );
  assert.equal(items[0].per, "pcs");
  assert.equal(items[0].quantity, 500);
  assert.equal(items[0].rate, 400);
});

test("the sample sheet parses without warnings", () => {
  const { warnings } = parseInvoiceSheet(sample);
  assert.deepEqual(warnings, []);
});

/** Build a minimal by-position invoice sheet, including the heading row. */
function sheetOf(items: SheetData, sameState: string | null = "no"): SheetData {
  return [
    ["labels"],
    ["TAX INVOICE", 1, "Buyer", "Address", 99, "GST", "01/01/2026", ""],
    ["labels"],
    ["", "", "", "", "", "", "", ""],
    ["labels"],
    ["Destination", "t1", "t2", "t3", sameState],
    ["SI", "Description", "HSN", "Qty", "Rate", "per", "GST", "Amount"],
    ...items,
  ];
}

test("reports the offending row and column for a non-numeric amount", () => {
  const { data, errors } = parseInvoiceSheet(
    sheetOf([[1, "Widget", 6307, 10, 5, "pcs", 5, "fifty"]]),
  );

  assert.equal(data, null);
  assert.equal(errors.length, 1);
  assert.match(errors[0].message, /Amount must be a number, but reads "fifty"/);
  assert.equal(errors[0].location, "row 8, column H (Amount)");
});

test("warns rather than fails when a GST rate is not a supported slab", () => {
  const { data, warnings } = parseInvoiceSheet(
    sheetOf([[1, "Widget", 6307, 10, 5, "pcs", 28, 50]]),
  );

  assert.ok(data);
  assert.equal(data.items[0].gstPercent, null);
  assert.equal(warnings.length, 1);
  assert.match(warnings[0].message, /GST % is 28, which is not a supported slab/);
  assert.equal(warnings[0].location, "row 8, column G (GST %)");
});

test("treats a same-state cell of 'yes' as intra-state, case-insensitively", () => {
  const item: SheetData = [[1, "Widget", 6307, 10, 5, "pcs", 5, 50]];
  assert.equal(parseInvoiceSheet(sheetOf(item, "Yes")).data!.meta.sameState, true);
  assert.equal(parseInvoiceSheet(sheetOf(item, "  YES ")).data!.meta.sameState, true);
  assert.equal(parseInvoiceSheet(sheetOf(item, "no")).data!.meta.sameState, false);
});

test("warns when the same-state cell is blank and falls back to inter-state", () => {
  const { data, warnings } = parseInvoiceSheet(
    sheetOf([[1, "Widget", 6307, 10, 5, "pcs", 5, 50]], null),
  );

  assert.equal(data!.meta.sameState, false);
  assert.equal(warnings.length, 1);
  assert.match(warnings[0].message, /Same-state flag is blank/);
});

test("skips blank rows between line items", () => {
  const { data } = parseInvoiceSheet(
    sheetOf([
      [1, "Widget", 6307, 10, 5, "pcs", 5, 50],
      [null, null, null, null, null, null, null, null],
      [2, "Gadget", 6307, 4, 5, "pcs", 5, 20],
    ]),
  );

  assert.equal(data!.items.length, 2);
  assert.deepEqual(data!.items.map((item) => item.description), ["Widget", "Gadget"]);
});

test("rejects an empty sheet with a readable message", () => {
  const { data, errors } = parseInvoiceSheet([]);
  assert.equal(data, null);
  assert.match(errors[0].message, /The sheet is empty/);
});

test("rejects a sheet that stops before the line-item rows", () => {
  const { data, errors } = parseInvoiceSheet(sheetOf([]));
  assert.equal(data, null);
  assert.match(errors[0].message, /needs at least 8/);
});
