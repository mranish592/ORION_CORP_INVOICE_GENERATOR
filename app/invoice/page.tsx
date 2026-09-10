import type { Metadata } from "next";

import { GeneratorClient } from "@/components/GeneratorClient";

export const metadata: Metadata = {
  title: "Tax invoice — Orion Corp",
};

export default function InvoicePage() {
  return (
    <>
      <p className="lede">
        Upload a filled-in invoice sheet to get a GST tax invoice PDF. The sheet is
        read by position, so keep every row where it is and change only the values.
      </p>

      <section className="panel">
        <p className="step-label">Start here</p>
        <p className="hint">
          <a href="/samples/invoice_sample.xlsx" download>
            Download the sample invoice sheet
          </a>{" "}
          and fill it in. Row 1 holds the labels; the values sit on rows 2, 4 and 6;
          the line items begin on row 8. Set the <code>Same State</code> cell to{" "}
          <code>yes</code> for a CGST + SGST split, or <code>no</code> for IGST. Every
          GST&nbsp;% must be <code>5</code>, <code>12</code> or <code>18</code>.
        </p>
      </section>

      <GeneratorClient kind="invoice" />

      <p className="privacy">
        The spreadsheet is parsed and the PDF is rendered in this browser tab. Nothing
        is uploaded and nothing is stored.
      </p>
    </>
  );
}
