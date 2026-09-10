# Orion Corp Invoice Generator

A Next.js app that turns an Excel spreadsheet into a formatted PDF. Two generators:

- **Tax invoice** — a GST invoice with line items, the CGST/SGST or IGST split, a tax
  summary and the amount in words.
- **Packing list** — one card per box, listing the products it contains.

**Everything runs in the browser.** The spreadsheet is parsed and the PDF is rendered
client-side. There is no database, no auth, no file storage and no upload endpoint: the
app is entirely statically prerendered, and the spreadsheet never leaves the device.
(The parser and pdfmake are loaded through `import()`, so the first upload and the first
generate each fetch a JavaScript chunk — those requests carry the app's own code, never
your data.) Deploys to Vercel with zero configuration.

This is a rewrite of a Django 3.0 + ReportLab app from 2020, which is preserved on the
[`legacy`](../../tree/legacy) branch. See `PORTING_PLAN.md` for the porting brief.

---

## Running it

```bash
npm install
npm run dev        # http://localhost:3000
```

| Script | What it does |
| --- | --- |
| `npm run dev` | Development server |
| `npm run build` | Production build |
| `npm start` | Serve the production build |
| `npm test` | Unit tests (`node --test`, no test-runner dependency) |
| `npm run typecheck` | `tsc --noEmit` |
| `npm run render:samples` | Render both sample sheets to `out/*.pdf` without a browser |
| `npm run build:signature` | Regenerate `lib/signatureImage.ts` from `public/sign.jpg` |

## Using it

Each generator page links to its sample sheet. Fill it in, choose the file, check the
figures the app reads back, then generate. The PDF opens in a preview and downloads as
a Blob — nothing is written to a server.

### Invoice sheet layout

The sheet is read **strictly by position**, so keep every row where it is and change
only the values. Row 1 holds column labels; values sit on rows 2, 4 and 6; line items
begin on row 8.

| Sheet row | Contents |
| --- | --- |
| 1 | Labels |
| **2** | Invoice type, invoice number, buyer name, buyer address, buyer phone, buyer GST, bill date, delivery note |
| 3 | Labels |
| **4** | Mode of payment, supplier ref, other ref, buyer order no., buyer order date, despatch document no., delivery note date, despatched through |
| 5 | Labels |
| **6** | Destination, term 1, term 2, term 3, same-state flag (`yes` / `no`) |
| 7 | Line-item headings |
| **8 onward** | One line item per row |

Line-item columns, in order: `SI`, `Description of Goods`, `HSN/SAC`, `Quantity`,
`Rate`, `per`, `GST %`, `Amount`.

- **GST %** must be `5`, `12` or `18`. Any other value still counts toward the taxable
  amount and the piece count but is not taxed, and the app raises a warning naming the
  row so the under-taxing is visible.
- **Same state** decides the tax split: `yes` prints CGST + SGST at half the rate each
  and a seven-column tax summary; anything else prints IGST and a five-column summary.

### Packing sheet layout

| Sheet row | Contents |
| --- | --- |
| 1 | Labels |
| **2** | Invoice number, shipping address |
| 3 | Labels |
| **4** | Buyer name, date |
| 5 | Label |
| **6** | Contact number |
| **7** | `Box` in column A, then one product name per column |
| **8 onward** | One row per box: box number in column A, quantities across |

A quantity of `0` (or a blank cell) omits that product from the box's card. Cards are
never split across a page break.

## How it is put together

```
app/                    Routes: /, /invoice, /packing
components/             Client components (file input, validation, preview)
config/company.ts       Orion Corp's own details — the only hard-coded content
lib/
  parseInvoiceSheet.ts  File -> typed object, with per-row/column validation
  parsePackingSheet.ts
  calculateTax.ts       Tax accumulation and the CGST/SGST vs IGST split
  amountInWords.ts      Indian lakh/crore wording
  buildInvoiceDocDefinition.ts   pdfmake document definitions
  buildPackingDocDefinition.ts
  pdfLayouts.ts         Shared table border layouts and the width budget
  pdfClient.ts          Lazy browser-only pdfmake loader
test/                   Unit tests against the sample sheets
```

Everything under `lib/` is pure and free of React and of the pdfmake runtime, which is
what makes it testable from Node.

### Dependencies

| Concern | Package | Why |
| --- | --- | --- |
| PDF generation | `pdfmake` | Declarative tables with `colSpan`/`rowSpan`, which the span-heavy GST invoice needs |
| Excel parsing | `read-excel-file` | TypeScript-native and browser-first |
| Number → words | `to-words` | Indian lakh/crore system (`en-IN` locale only) |

**Do not `npm install xlsx`.** SheetJS is frozen on the public npm registry at 0.18.5
with two unfixed high-severity advisories
([GHSA-4r6h-8v6p-xvw6](https://github.com/advisories/GHSA-4r6h-8v6p-xvw6),
[GHSA-5pgg-2g8v-p4x9](https://github.com/advisories/GHSA-5pgg-2g8v-p4x9)) and no upgrade
path — the fixed versions were moved to SheetJS's own CDN. `npm audit` currently reports
zero advisories; keep it that way.

### Two things worth knowing

- **pdfmake adds cell padding and rule widths *outside* the declared column widths.**
  Feeding it the legacy 540pt figures makes the table wider than an A4 page and the
  right-hand columns silently wrap or fall off it. `contentWidths()` in
  `lib/pdfLayouts.ts` converts the legacy full-column widths into pdfmake content
  widths; use it for every table.
- **pdfmake must be imported lazily from a client component.** It touches browser
  globals on evaluation and will break during prerender otherwise. `lib/pdfClient.ts`
  handles this.

## Changes from the legacy app

Confirmed with the business during the port:

- The **ICICI** bank block is the current one. The legacy source also built a Punjab &
  Sind Bank block and then discarded it without rendering; that dead path is gone.
- Figures now print the **₹** symbol, which pdfmake's bundled Roboto supports and
  ReportLab's Helvetica did not. The amount in words keeps the legally conventional
  `INR … Rupees and … Paisa Only` wording.
- The signature line reads **`for Orion Corp`**, correcting the legacy
  `for Orion the Corp`.
- The **by-position sheet format is unchanged**, so every sheet the business already
  has keeps working. Parsing is more tolerant (blank rows and trailing blank columns
  are ignored) and every problem is reported against a sheet row and column.

Bugs found in the legacy source and fixed here:

1. HSN/SAC codes were Python `set` objects placed straight into a cell, so the PDF
   printed `{6307}`. They are now joined as a comma-separated list.
2. The tax-amount-in-words line appended `' only'` to a string already ending in
   `' Paisa Only'`, printing `… Paisa Only only`. One `Only` now.
3. Totals accumulated in module-level globals that two concurrent requests could
   corrupt. All state is local to a call.
4. Both generators wrote to a single fixed path in `static/`, so concurrent users
   overwrote each other's PDFs and "Preview" could show someone else's invoice. The PDF
   is now a Blob handed straight to the user and never persisted.
5. `'Buyer Name: ' + buyer_name` crashed on a numeric cell. Every cell is coerced to
   text.
6. A malformed sheet produced an opaque traceback. Sheets are validated up front and
   errors name the offending row and column, e.g.
   *row 8, column G (GST %) — GST % is 28, which is not a supported slab.*
   (`read-excel-file`'s schema support assumes a header row followed by value rows, so
   it cannot describe this by-position sheet; the equivalent validation is hand-written
   in the parsers.)
7. The `for Orion the Corp` typo, corrected as above.

Two other differences, both deliberate:

- Tax slabs are emitted in ascending order (5, 12, 18). The legacy code happened to
  emit them 5, 18, 12.
- A literal `\n` typed into an address cell becomes a real line break instead of being
  swallowed.
