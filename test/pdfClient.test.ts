import assert from "node:assert/strict";
import { test } from "node:test";

import { buildInvoiceDocDefinition } from "../lib/buildInvoiceDocDefinition";
import { buildPackingDocDefinition } from "../lib/buildPackingDocDefinition";
import { calculateInvoiceTotals } from "../lib/calculateTax";
import { parseInvoiceSheet } from "../lib/parseInvoiceSheet";
import { parsePackingSheet } from "../lib/parsePackingSheet";
import { renderPdfBlob } from "../lib/pdfClient";
import { SIGNATURE_DATA_URL } from "../lib/signatureImage";
import { readSample } from "./helpers";

/**
 * These drive `renderPdfBlob` — the same lazy-loading path the browser takes,
 * against the same pdfmake browser bundle. pdfmake 0.3 returns a Promise from
 * `getBlob()` where 0.2 took a callback; calling it the 0.2 way hangs forever
 * instead of failing, so it is worth pinning down.
 */

async function pdfHeader(blob: Blob): Promise<string> {
  return Buffer.from(await blob.slice(0, 5).arrayBuffer()).toString("latin1");
}

test("renders an invoice to a PDF blob through the browser pdfmake build", async () => {
  const sheet = parseInvoiceSheet(await readSample("invoice_sample.xlsx")).data!;
  const blob = await renderPdfBlob(
    buildInvoiceDocDefinition({
      sheet,
      totals: calculateInvoiceTotals(sheet.items, sheet.meta.sameState),
      signature: SIGNATURE_DATA_URL,
    }),
  );

  assert.equal(blob.type, "application/pdf");
  assert.ok(blob.size > 10_000);
  assert.equal(await pdfHeader(blob), "%PDF-");
});

test("renders a packing list to a PDF blob", async () => {
  const sheet = parsePackingSheet(await readSample("packing_sample.xlsx")).data!;
  const blob = await renderPdfBlob(buildPackingDocDefinition(sheet));

  assert.equal(blob.type, "application/pdf");
  assert.ok(blob.size > 1_000);
  assert.equal(await pdfHeader(blob), "%PDF-");
});
