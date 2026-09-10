import assert from "node:assert/strict";
import { test } from "node:test";

import { amountInWords } from "../lib/amountInWords";

test("matches the amount printed on the reference invoice", () => {
  assert.equal(
    amountInWords(355605),
    "INR Three Lakh Fifty Five Thousand Six Hundred And Five Rupees and Zero Paisa Only",
  );
});

test("matches the tax amount printed on the reference invoice", () => {
  assert.equal(
    amountInWords(23805),
    "INR Twenty Three Thousand Eight Hundred And Five Rupees and Zero Paisa Only",
  );
});

test("restores the 'and' before a remainder under one hundred", () => {
  assert.equal(amountInWords(452).startsWith("INR Four Hundred And Fifty Two Rupees"), true);
  assert.equal(amountInWords(1005).startsWith("INR One Thousand And Five Rupees"), true);
  assert.equal(amountInWords(105).startsWith("INR One Hundred And Five Rupees"), true);
});

test("omits the 'and' when there is no remainder under one hundred", () => {
  assert.equal(amountInWords(276300).startsWith("INR Two Lakh Seventy Six Thousand Three Hundred Rupees"), true);
  assert.equal(amountInWords(100).startsWith("INR One Hundred Rupees"), true);
});

test("does not use 'and' for numbers below one hundred", () => {
  assert.equal(amountInWords(21), "INR Twenty One Rupees and Zero Paisa Only");
});

test("uses the Indian lakh and crore system", () => {
  assert.equal(amountInWords(10000000).startsWith("INR One Crore Rupees"), true);
  assert.equal(amountInWords(1000000).startsWith("INR Ten Lakh Rupees"), true);
});

test("renders paisa and keeps the legacy 'Paisa Only' wording", () => {
  assert.equal(
    amountInWords(452.36),
    "INR Four Hundred And Fifty Two Rupees and Thirty Six Paisa Only",
  );
});

test("never rolls the paisa over to one hundred", () => {
  assert.equal(amountInWords(1.999), "INR Two Rupees and Zero Paisa Only");
  assert.equal(amountInWords(0.005), "INR Zero Rupees and One Paisa Only");
});

test("handles zero", () => {
  assert.equal(amountInWords(0), "INR Zero Rupees and Zero Paisa Only");
});
