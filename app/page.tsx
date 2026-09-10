import Link from "next/link";

export default function HomePage() {
  return (
    <>
      <p className="lede">
        Upload a spreadsheet and get a formatted PDF back. Everything runs in your
        browser — the file is never uploaded, stored, or sent anywhere.
      </p>

      <div className="card-grid">
        <Link className="card" href="/invoice">
          <h2>Tax invoice →</h2>
          <p>
            A single GST invoice with the line items, the CGST/SGST or IGST split,
            the tax summary and the amount in words.
          </p>
        </Link>
        <Link className="card" href="/packing">
          <h2>Packing list →</h2>
          <p>
            One card per box listing the products it contains, paginated so a card
            is never split across pages.
          </p>
        </Link>
      </div>

      <p className="privacy">
        No account, no database, no upload endpoint. Your spreadsheet is read and the
        PDF is built on this device; the only requests this page makes are for its own
        code.
      </p>
    </>
  );
}
