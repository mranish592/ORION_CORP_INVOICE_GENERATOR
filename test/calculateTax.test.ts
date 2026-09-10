import assert from "node:assert/strict";
import { test } from "node:test";

import { calculateInvoiceTotals, taxRowsFor } from "../lib/calculateTax";
import { parseInvoiceSheet } from "../lib/parseInvoiceSheet";
import type { LineItem } from "../lib/types";
import { readSample } from "./helpers";

const sample = await readSample("invoice_sample.xlsx");
const parsedSample = parseInvoiceSheet(sample).data!;

function item(partial: Partial<LineItem>): LineItem {
  return {
    si: "1",
    description: "Widget",
    hsn: "6307",
    quantity: 1,
    rate: 1,
    per: "pcs",
    gstPercent: 5,
    amount: 100,
    ...partial,
  };
}

test("reproduces every figure on the reference invoice", () => {
  const totals = calculateInvoiceTotals(parsedSample.items, parsedSample.meta.sameState);

  assert.equal(totals.taxableAmount, 331800);
  assert.equal(totals.totalPieces, 2700);
  assert.equal(totals.totalTax, 23805);
  assert.equal(totals.totalAmount, 355605);

  assert.deepEqual(
    totals.slabs.map((slab) => [slab.rate, slab.tax, slab.taxableValue]),
    [
      [5, 13815, 276300],
      [18, 9990, 55500],
    ],
  );
});

test("collects the distinct HSN codes per slab", () => {
  const totals = calculateInvoiceTotals(parsedSample.items, false);
  assert.deepEqual(totals.slabs[0].hsnCodes, ["6307"]);
  assert.deepEqual(totals.slabs[1].hsnCodes, ["3921", "40151900", "6307"]);
});

test("skips no slab that carries tax and omits slabs that carry none", () => {
  const totals = calculateInvoiceTotals(
    [item({ gstPercent: 12, amount: 1000 }), item({ gstPercent: 18, amount: 1000 })],
    false,
  );
  assert.deepEqual(totals.slabs.map((slab) => slab.rate), [12, 18]);
});

test("a row outside the supported slabs is still counted but never taxed", () => {
  const totals = calculateInvoiceTotals(
    [item({ gstPercent: null, amount: 500, quantity: 7 })],
    false,
  );

  assert.equal(totals.taxableAmount, 500);
  assert.equal(totals.totalPieces, 7);
  assert.equal(totals.totalTax, 0);
  assert.equal(totals.totalAmount, 500);
  assert.deepEqual(totals.slabs, []);
});

test("rounds once after summation rather than per row", () => {
  // Each row's 5% tax is 0.0055; summed first this is 0.011 -> 0.01, whereas
  // rounding per row would have produced 0.02.
  const totals = calculateInvoiceTotals(
    [item({ amount: 0.11 }), item({ amount: 0.11 })],
    false,
  );
  assert.equal(totals.slabs[0].tax, 0.01);
});

test("splits each slab into equal CGST and SGST halves when intra-state", () => {
  const totals = calculateInvoiceTotals(parsedSample.items, true);
  assert.deepEqual(taxRowsFor(totals), [
    { label: "CGST 2.5%", amount: 6907.5 },
    { label: "SGST 2.5%", amount: 6907.5 },
    { label: "CGST 9%", amount: 4995 },
    { label: "SGST 9%", amount: 4995 },
  ]);
});

test("charges a single IGST row per slab when inter-state", () => {
  const totals = calculateInvoiceTotals(parsedSample.items, false);
  assert.deepEqual(taxRowsFor(totals), [
    { label: "IGST 5%", amount: 13815 },
    { label: "IGST 18%", amount: 9990 },
  ]);
});

test("labels the 12% intra-state split as CGST/SGST 6%", () => {
  const totals = calculateInvoiceTotals([item({ gstPercent: 12, amount: 1000 })], true);
  assert.deepEqual(taxRowsFor(totals), [
    { label: "CGST 6%", amount: 60 },
    { label: "SGST 6%", amount: 60 },
  ]);
});

test("keeps totals local to the call", () => {
  // Legacy bug §7.3: totals accumulated in module-level globals.
  const first = calculateInvoiceTotals(parsedSample.items, false);
  const second = calculateInvoiceTotals(parsedSample.items, false);
  assert.deepEqual(first, second);
});
