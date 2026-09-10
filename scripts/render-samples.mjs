/**
 * Renders both sample sheets to PDFs in `out/` without a browser, so layout
 * changes can be eyeballed and regressions caught from the command line.
 *
 * The app itself never runs pdfmake on a server — this is a development aid.
 * Run with: npm run render:samples
 */
import { mkdirSync, writeFileSync } from "node:fs";
import { createRequire } from "node:module";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";

import { readSheet } from "read-excel-file/node";

import { buildInvoiceDocDefinition } from "../lib/buildInvoiceDocDefinition.ts";
import { buildPackingDocDefinition } from "../lib/buildPackingDocDefinition.ts";
import { calculateInvoiceTotals } from "../lib/calculateTax.ts";
import { parseInvoiceSheet } from "../lib/parseInvoiceSheet.ts";
import { parsePackingSheet } from "../lib/parsePackingSheet.ts";
import { SIGNATURE_DATA_URL } from "../lib/signatureImage.ts";

const require = createRequire(import.meta.url);
const pdfMake = require("pdfmake");

const fontDir = join(dirname(require.resolve("pdfmake/package.json")), "build/fonts/Roboto");
pdfMake.setFonts({
  Roboto: {
    normal: join(fontDir, "Roboto-Regular.ttf"),
    bold: join(fontDir, "Roboto-Medium.ttf"),
    italics: join(fontDir, "Roboto-Italic.ttf"),
    bolditalics: join(fontDir, "Roboto-MediumItalic.ttf"),
  },
});
pdfMake.setUrlAccessPolicy(() => false);
pdfMake.setLocalAccessPolicy((path) => path.startsWith(fontDir));

const outDir = fileURLToPath(new URL("../out", import.meta.url));
mkdirSync(outDir, { recursive: true });

function report(label, { errors, warnings }) {
  for (const issue of errors) {
    console.error(`  error   ${label}: ${issue.location ?? ""} ${issue.message}`);
  }
  for (const issue of warnings) {
    console.warn(`  warning ${label}: ${issue.location ?? ""} ${issue.message}`);
  }
  if (errors.length > 0) process.exitCode = 1;
}

async function write(name, docDefinition) {
  const buffer = await pdfMake.createPdf(docDefinition).getBuffer();
  const target = join(outDir, name);
  writeFileSync(target, buffer);
  console.log(`  wrote   ${target} (${buffer.length} bytes)`);
}

const invoice = parseInvoiceSheet(await readSheet("public/samples/invoice_sample.xlsx"));
report("invoice", invoice);

if (invoice.data) {
  for (const sameState of [false, true]) {
    const sheet = { ...invoice.data, meta: { ...invoice.data.meta, sameState } };
    const totals = calculateInvoiceTotals(sheet.items, sameState);
    console.log(
      `  totals  ${sameState ? "intra-state" : "inter-state"}: taxable ${totals.taxableAmount}, ` +
        `tax ${totals.totalTax}, total ${totals.totalAmount}, pieces ${totals.totalPieces}`,
    );
    await write(
      sameState ? "invoice-same-state.pdf" : "invoice-inter-state.pdf",
      buildInvoiceDocDefinition({ sheet, totals, signature: SIGNATURE_DATA_URL }),
    );
  }
}

const packing = parsePackingSheet(await readSheet("public/samples/packing_sample.xlsx"));
report("packing", packing);
if (packing.data) await write("packing.pdf", buildPackingDocDefinition(packing.data));
