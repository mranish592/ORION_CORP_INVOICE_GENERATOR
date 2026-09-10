import type { Metadata } from "next";

import { GeneratorClient } from "@/components/GeneratorClient";

export const metadata: Metadata = {
  title: "Packing list — Orion Corp",
};

export default function PackingPage() {
  return (
    <>
      <p className="lede">
        Upload a packing sheet to get one card per box. Products with a quantity of 0
        are left off that box&rsquo;s card, and no card is ever split across a page break.
      </p>

      <section className="panel">
        <p className="step-label">Start here</p>
        <p className="hint">
          <a href="/samples/packing_sample.xlsx" download>
            Download the sample packing sheet
          </a>{" "}
          and fill it in. Row 1 holds the labels; the values sit on rows 2, 4 and 6;
          row 7 names the products across the columns, and the box rows begin on row 8
          with the box number in column&nbsp;A.
        </p>
      </section>

      <GeneratorClient kind="packing" />

      <p className="privacy">
        The spreadsheet is parsed and the PDF is rendered in this browser tab. Nothing
        is uploaded and nothing is stored.
      </p>
    </>
  );
}
