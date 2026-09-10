import assert from "node:assert/strict";
import { test } from "node:test";

import { parsePackingSheet } from "../lib/parsePackingSheet";
import type { SheetData } from "../lib/types";
import { readSample } from "./helpers";

const sample = await readSample("packing_sample.xlsx");

test("reads the metadata rows by position", () => {
  const { data, errors } = parsePackingSheet(sample);
  assert.deepEqual(errors, []);
  assert.ok(data);

  assert.equal(data.invoiceNumber, "72");
  assert.equal(data.buyerName, "SOUTHERN RAILWAY");
  assert.equal(data.date, "31/07/2020");
  assert.equal(data.contactNumber, "9003840110");
  assert.match(data.shippingAddress, /^OFFICE OF THE CWM CHIEF WORKSHOP MANAGER ,/);
});

test("turns the literal \\n typed into the address cell into real line breaks", () => {
  const { data } = parsePackingSheet(sample);
  assert.equal(data!.shippingAddress.includes("\\n"), false);
  assert.equal(data!.shippingAddress.split("\n").length, 3);
});

test("emits one box per row, omitting zero and blank quantities", () => {
  const { data } = parsePackingSheet(sample);
  const boxes = data!.boxes;

  assert.equal(boxes.length, 5);
  assert.deepEqual(boxes.map((box) => box.boxNumber), ["1", "2", "3", "4", "5"]);

  assert.deepEqual(boxes[0].lines, [
    { product: "Blue Respirator", quantity: 300 },
    { product: "Pink respirator", quantity: 200 },
  ]);
  // Box 2 has an explicit 0 for "Pink respirator", which is left off the card.
  assert.deepEqual(boxes[1].lines, [{ product: "Blue Respirator", quantity: 100 }]);
  assert.deepEqual(boxes[4].lines, [
    { product: "Pink respirator", quantity: 100 },
    { product: "White", quantity: 400 },
    { product: "Pink", quantity: 800 },
    { product: "Blue", quantity: 400 },
  ]);
});

test("the sample sheet parses without warnings", () => {
  assert.deepEqual(parsePackingSheet(sample).warnings, []);
});

function sheetOf(rows: SheetData): SheetData {
  return [
    ["labels"],
    [7, "Ship to"],
    ["labels"],
    ["Buyer", "01/01/2026"],
    ["Contact No:"],
    [90000],
    ["Box", "Alpha", "Beta"],
    ...rows,
  ];
}

test("reports the offending row and column for a non-numeric quantity", () => {
  const { data, errors } = parsePackingSheet(sheetOf([[1, "many", 2]]));

  assert.equal(data, null);
  assert.equal(errors.length, 1);
  assert.match(errors[0].message, /Quantity for "Alpha" must be a number, but reads "many"/);
  assert.equal(errors[0].location, "row 8, column B (Alpha)");
});

test("warns about a box whose products are all zero", () => {
  const { data, warnings } = parsePackingSheet(sheetOf([[4, 0, 0]]));

  assert.equal(data!.boxes.length, 1);
  assert.deepEqual(data!.boxes[0].lines, []);
  assert.match(warnings[0].message, /Box 4 has no products with a quantity above 0/);
});

test("ignores columns with no product name", () => {
  const rows = sheetOf([[1, 5, 6]]);
  rows[6] = ["Box", "Alpha", null, "Gamma"];
  const { data } = parsePackingSheet(rows);
  assert.deepEqual(data!.boxes[0].lines, [{ product: "Alpha", quantity: 5 }]);
});

test("rejects a sheet with no product names", () => {
  const rows = sheetOf([[1, 5, 6]]);
  rows[6] = ["Box"];
  const { data, errors } = parsePackingSheet(rows);
  assert.equal(data, null);
  assert.match(errors[0].message, /No product names found/);
  assert.equal(errors[0].location, "row 7, column B (Product names)");
});

test("rejects an empty sheet with a readable message", () => {
  const { data, errors } = parsePackingSheet([]);
  assert.equal(data, null);
  assert.match(errors[0].message, /The sheet is empty/);
});
