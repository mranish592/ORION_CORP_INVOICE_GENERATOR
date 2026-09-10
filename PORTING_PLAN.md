# Orion Corp Invoice Generator — Port to Next.js

**You are starting from an empty repository.** Everything you need is either in this
document or on the `legacy` git branch. Read this file end to end before writing code.

---

## 1. What you are building

A **Vercel-deployable Next.js (TypeScript) app** that turns an uploaded Excel
spreadsheet into a formatted PDF. Two independent generators:

| Generator | Input | Output |
| --- | --- | --- |
| **Tax invoice** | `invoice_sample.xlsx` shape | Single GST invoice PDF (India) |
| **Packing list** | `packing_sample.xlsx` shape | One card per box, paginated |

**This is a client-side app.** The spreadsheet is parsed and the PDF is rendered in
the browser. There is no database, no auth, no file storage, and no upload endpoint.
The user's file never leaves their machine. Do not add a backend unless something
below explicitly requires one — nothing does.

### Origin

The predecessor is a Django 3.0 + ReportLab app written in 2020. It works but is
six years old. The port is a **rewrite, not a translation** — reuse the *business
rules and layout*, not the code structure.

---

## 2. Where the reference implementation is

The old app is preserved on the **`legacy`** branch of this repository. It is
verified working. Read it directly:

```bash
git show legacy:orion_website/invoice_generator/generate_pdf.py        # 518 lines — invoice layout
git show legacy:orion_website/packing_list/generate_packing_list.py    #  87 lines — packing layout
git show legacy:orion_website/invoice_generator/views.py               # how the sheet is parsed
git show legacy:README.md                                              # sheet format notes
```

Binary assets you need — extract these into the new project:

```bash
git show legacy:orion_website/static/invoice_sample.xlsx > public/samples/invoice_sample.xlsx
git show legacy:orion_website/static/packing_sample.xlsx > public/samples/packing_sample.xlsx
git show legacy:orion_website/static/images/sign.jpg     > public/sign.jpg
```

There are also reference PDFs showing the exact target output. Extract and open them —
**your output should look like these**:

```bash
git show legacy:orion_website/static/invoice_download.pdf > /tmp/reference_invoice.pdf
git show legacy:orion_website/static/packing_download.pdf > /tmp/reference_packing.pdf
```

To run the old app for side-by-side comparison, check out `legacy` in a worktree and
follow its README (Python 3.12 venv + `requirements.txt`).

---

## 3. Stack

These were chosen after checking the npm registry and advisory database. Use them
unless you find a concrete blocker.

| Concern | Package | Version at time of writing |
| --- | --- | --- |
| PDF generation | **`pdfmake`** | 0.3.11 |
| Excel parsing | **`read-excel-file`** | 9.3.10 |
| Number → words (Indian) | **`to-words`** | 7.1.0 |

### Why pdfmake

pdfmake takes a declarative document-definition object — tables, column widths,
styles — which maps almost 1:1 onto ReportLab's platypus model that the legacy code
uses. The legacy invoice relies on **merged cells 14 times**; pdfmake has native
`colSpan`/`rowSpan`, plus `headerRows` for repeating headers, custom border layouts
via `hLineWidth`/`vLineWidth`, `alignment`, and fixed `widths`.

Alternatives considered and rejected:

- **`@react-pdf/renderer`** — better maintained and nicer React DX, but it has **no
  Table component**. Its primitives are `Document, Page, View, Image, Text, Link,
  Note, Canvas`. Every merged cell would be hand-rolled flexbox width math. Bad fit
  for a span-heavy GST invoice.
- **`jspdf` + `jspdf-autotable`** — autotable handles the line-item grid, but the
  surrounding header/tax/declaration blocks become imperative `doc.text(x, y)`
  positioning. A step backwards.
- **`pdf-lib`** — last published November 2021 and has no layout engine.
- **Puppeteer HTML→PDF** — server-only, needs `@sparticuz/chromium` on serverless,
  defeats the client-side goal.

### ⚠️ Do NOT `npm install xlsx`

The SheetJS `xlsx` package on the public npm registry is **frozen at 0.18.5 (March
2022)** and carries two unfixed high-severity advisories:

- [GHSA-4r6h-8v6p-xvw6](https://github.com/advisories/GHSA-4r6h-8v6p-xvw6) — Prototype Pollution, CVSS 7.8, fixed in `<0.19.3`
- [GHSA-5pgg-2g8v-p4x9](https://github.com/advisories/GHSA-5pgg-2g8v-p4x9) — ReDoS, CVSS 7.5, fixed in `<0.20.2`

**Those fixed versions do not exist on npm** — SheetJS moved distribution to their
own CDN. `npm audit` will flag it permanently with no upgrade path. Use
`read-excel-file` instead; it is TypeScript-native, browser-first, and has a schema
validation feature you should use (see §6).

### Next.js notes

- pdfmake touches browser globals. Import it lazily inside a `"use client"`
  component (`const pdfMake = (await import("pdfmake/build/pdfmake")).default`) or
  via `next/dynamic` with `ssr: false`. It **will** break during SSR/prerender otherwise.
- pdfmake ships fonts through a VFS file (`vfs_fonts.js`, Roboto). Custom fonts need
  a VFS build step.
- **Currency symbol:** the legacy PDF prints the literal string `INR` because
  ReportLab's default Helvetica has no `₹` glyph. Roboto *does* have `₹`. Using the
  real symbol is an improvement — but keep the words `INR ... Rupees ... Paisa Only`
  in the amount-in-words line exactly as specified in §5.3, since that is legally
  conventional wording on Indian invoices.
- Import only the `en-IN` locale from `to-words`; the full package carries 136 locales.

---

## 4. Input format — invoice sheet

**Critical:** the legacy parser reads the sheet **strictly by row and column index**.
Metadata is interleaved as alternating label/value rows. Open
`public/samples/invoice_sample.xlsx` before you write the parser.

Structure of the parsed rows (0-indexed, *after* the spreadsheet's own header row is
consumed as column names):

| Row index | Contents |
| --- | --- |
| **0** | **Values:** invoice type, invoice number, buyer name, buyer address, buyer phone, buyer GST, bill date, delivery note |
| 1 | Labels (ignored) |
| **2** | **Values:** mode of payment, supplier ref, other ref, buyer order no., buyer order date, despatch document no., delivery note date, despatched through |
| 3 | Labels (ignored) |
| **4** | **Values:** destination, term1, term2, term3, same-state flag (`yes`/`no`) |
| 5 | Line-item column headings (ignored — replaced with canonical headings) |
| **6 onward** | **One line item per row** |

The label rows exist purely so a human filling the sheet knows what goes underneath.
The legacy code discards rows 0–5 with six `pop(0)` calls, then prepends its own
header row.

Line-item columns, in order:

| Idx | Column | Type | Used for |
| --- | --- | --- | --- |
| 0 | SI | number | Serial number, displayed only |
| 1 | Description of Goods | string | Wraps inside the cell (font size 6, centered) |
| 2 | HSN/SAC | number | Grouped per tax slab in the tax summary |
| 3 | Quantity | number | Summed → total pieces |
| 4 | Rate | number | Displayed only |
| 5 | per | string | Unit, e.g. `pcs` |
| 6 | GST % | number | **Must be exactly `5`, `12`, or `18`** |
| 7 | Amount | number | Summed → taxable value |

Blank cells are coerced to `''` (empty string) before parsing.

## 4b. Input format — packing sheet

| Row index | Contents |
| --- | --- |
| **0** | **Values:** invoice number, shipping address |
| 1 | Labels (ignored) |
| **2** | **Values:** buyer name, date |
| 3 | Labels (ignored) |
| **4** | **Values:** contact number |
| **5** | **Product names** — column 0 is `Box`, columns 1..n are product names |
| **6 onward** | One row per box: column 0 = box number, columns 1..n = quantity of that product |

Blank cells are coerced to `0`. **A quantity of `0` means "omit this product from
this box's card"** — it is not printed as a zero row.

---

## 5. Business logic (port this exactly)

### 5.1 Tax accumulation

Iterate the line items. For each row, using GST% (col 6) and Amount (col 7):

```
amount        += row[7]                       // running taxable total
total_pieces  += row[3]

if row[6] == 5   -> tax_5  += 0.05 * row[7];  tax_5_set.add(row[2])
if row[6] == 12  -> tax_12 += 0.12 * row[7];  tax_12_set.add(row[2])
if row[6] == 18  -> tax_18 += 0.18 * row[7];  tax_18_set.add(row[2])
```

`tax_N_set` is the **set of distinct HSN/SAC codes** in that slab — it becomes the
first cell of the tax-summary row.

```
total_amount = amount + tax_5 + tax_12 + tax_18
```

Round `tax_5`, `tax_12`, `tax_18`, `amount`, `total_amount` to 2 decimals **after**
summation, not per row.

A row whose GST% is not exactly 5, 12 or 18 still contributes to `amount` and
`total_pieces` but generates no tax. Preserve this behaviour, but see §7 — you should
surface a validation warning rather than silently under-taxing.

### 5.2 Same-state split (CGST/SGST vs IGST)

Driven by the same-state flag at row 4, column 4, lowercased and compared to `'yes'`:

- **`yes`** (intra-state) → each slab splits into two half-rate rows:
  5% → `CGST 2.5%` + `SGST 2.5%`, each `tax_5 / 2`
  12% → `CGST 6%` + `SGST 6%`, each `tax_12 / 2`
  18% → `CGST 9%` + `SGST 9%`, each `tax_18 / 2`
- **anything else** (inter-state) → one row per slab: `IGST 5%` / `IGST 12%` /
  `IGST 18%` at the full slab amount.

This also changes the **tax summary table's column layout** — see §6.3.

### 5.3 Amount in words

Legacy uses Python `num2words(n, lang='en_IN')`, producing the Indian
lakh/crore system. Replace with `to-words`:

```ts
import { ToWords } from "to-words";
const toWords = new ToWords({ localeCode: "en-IN" });
```

The legacy output string is built as:

```
'INR ' + <rupees in words, Title Case> + ' Rupees and ' + <paisa in words, Title Case> + ' Paisa Only'
```

where `rupees = Math.trunc(n)` and `paisa = Math.round((n - rupees) * 100)`, and any
hyphens in the words are replaced with spaces and commas removed. Reproduce this
exact wording. Verify against the reference PDF.

`to-words` can produce the currency phrasing directly
(`toWords(452.36, { currency: true })` → `"Four Hundred Fifty Two Rupees And Thirty
Six Paise Only"`), but note it says **"Paise"** where the legacy says **"Paisa"**, and
omits the leading `INR`. Match the legacy string.

---

## 6. Output format — invoice PDF

Page **A4**, zero page margins, content width **540pt** (ReportLab centres the tables,
giving ~27pt effective side margins). Default font size 6 for line items.

Stacked vertically, with no gaps between blocks (they share borders):

```
Spacer(20)
<invoice type>            centred, font size 12   <- from row 0 col 0, e.g. "PROFORMA INVOICE"
Spacer(20)
[1] fixed details table
[2] buyer + terms table
[3] line items table
[4] amount chargeable table
[5] tax summary table
[6] declaration + signature table
Spacer(6)
"This is a Computer Generated Invoice"   centred, font size 10
```

### 6.1 Fixed details table — widths `[280, 130, 130]`

Six rows. Column 0 is **one merged cell spanning all six rows** containing the
hard-coded Orion Corp company block (name, address, warehouse, GSTIN, bank details) as
rich text at font size 8. Columns 1–2 carry paired label/value cells:

| Row | Col 1 | Col 2 |
| --- | --- | --- |
| 0 | `Inv No.` + invoice number | `Dated` + bill date |
| 1 | `Delivery Note` + note | `Mode/Terms of Payment` |
| 2 | `Other Reference(s)` | `Buyer's Order No.` |
| 3 | `Order Date:` | `Despatch Document No.` |
| 4 | `Delivery Note Date` | `Despatch through` |
| 5 | `SHIP TO:` + destination — **spans cols 1–2** | (merged) |

Full grid borders, 0.5pt grey. All cells top-aligned.

The company block is **hard-coded in the legacy source**, not read from the sheet.
Copy it verbatim from `generate_pdf.py` lines 139–150 — it includes the GSTIN, bank
account and IFSC. Put it in a single `config/company.ts` constant so it is editable
in one place.

### 6.2 Buyer + terms table — widths `[280, 260]`

Four rows, two columns:

| Row | Col 0 | Col 1 |
| --- | --- | --- |
| 0 | `Buyer Name: ` + name | `Terms of Delivery` |
| 1 | `Address: ` + address | `*` + term1 |
| 2 | `Phone no.: ` + phone | `*` + term2 |
| 3 | `GST/PAN no.: ` + GST | `*` + term3 |

Borders: grid on row 0 only, outer box around the whole table, and a box around
column 0 rows 1–3. Top-aligned, justified text at font size 8.

### 6.3 Line items table — widths `[20, 210, 60, 60, 50, 30, 40, 70]`

Header row: `SI | Description of Goods | HSN/SAC | Quantity | Rate | per | GST% | Amount`

Then the line items, then appended summary rows:

1. Subtotal row — only col 7 populated with `amount`
2. The tax rows from §5.2 (label in col 1, `%` in col 5, value in col 7)
3. `Total` row — col 1 `Total`, col 3 `total_pieces`, col 5 `%`, col 7 `total_amount`

Borders: full grid over the header + line items, a box around the summary block, a
grid on the final Total row, and vertical rules between all columns through the
summary region. The legacy code computes these boundaries with negative indices
scaled by the number of active tax slabs (`flag = flag_5 + flag_12 + flag_18`, doubled
when same-state). **In pdfmake, express this with a custom `layout` whose
`hLineWidth`/`vLineWidth` functions test the row index against the summary start —
much clearer than reproducing the arithmetic.**

Descriptions render as wrapped, centred text at font size 6.

### 6.4 Amount chargeable table — widths `[470, 70]`

| Row | Content |
| --- | --- |
| 0 | `Amount chargeable (inwords)` \| `E. &O.E` (right-aligned) |
| 1 | the amount-in-words string, **spanning both columns** |

Outer box only.

### 6.5 Tax summary table — layout depends on same-state

**Same state (`yes`)** — widths `[120, 80, 40, 80, 40, 80, 100]`:

```
| HSN/SAC | Taxable Value | Central Tax   | State Tax     | Total Tax Amount |
|         |               | Rate | Amount | Rate | Amount |                  |
```

Row 0 cells `HSN/SAC`, `Taxable Value` and `Total Tax Amount` each span both header
rows; `Central Tax` spans cols 2–3; `State Tax` spans cols 4–5.

One body row per active slab: the HSN set, the back-computed taxable value
(`tax_N / rate`, rounded to 2dp), rate label, half the tax, rate label, half the tax,
full slab tax. Then a `Total` row: `amount`, half total tax, half total tax, total tax.

**Inter-state** — widths `[240, 80, 40, 80, 100]`:

```
| HSN/SAC | Taxable Value | Integrated Tax | Total Tax Amount |
|         |               | Rate | Amount  |                  |
```

Same idea with one rate/amount pair. Full grid, top-aligned.

### 6.6 Declaration + signature table — widths `[280, 260]`, row 0 height 20

| Row | Col 0 | Col 1 |
| --- | --- | --- |
| 0 | `Tax Amount (in words):` + words — **spans both cols** | (merged) |
| 1 | `Declaration` | `for Orion the Corp` (right-aligned) |
| 2 | MSME number + GeM Seller ID | **signature image** `sign.jpg`, 100×60 |
| 3 | "We declare that this invoice shows the actual price of the goods described and that all particulars are true and correct." | `Authorised Signatory` (right-aligned) |

Embed the signature as a base64 data URI in pdfmake's `images` map.

---

## 6b. Output format — packing list PDF

Page A4, zero margins, `Spacer(40)` at the top. **One card per box**, each a
two-column table with widths `[300, 240]`:

```
Box No: <n>                                       <- spans, no border
Invoice no: <inv>                | Date: <date>   <- right-aligned col 1
Buyer: <name>                    | Contact No: <contact>
Address: <shipping address>
(blank row)
Product                          | Quantity       <- grid borders start here
<product name>                   | <qty>          <- one row per NON-ZERO quantity
...
```

Borders apply only from the `Product | Quantity` header row down. `Spacer(40)` after
each card.

**Pagination:** the legacy code uses a crude manual budget — `spaceleft` starts at 42,
each card subtracts `10 + address_line + x` (where `address_line = len(address) / 50`
and `x` = number of non-zero products), and when `spaceleft < 2` it emits a page break
and resets. **Do not port this.** Use pdfmake's natural flow with
`unbreakable: true` on each card table so cards never split across pages.

---

## 7. Legacy bugs — do NOT carry these over

Found while reading the source. Fix them in the port.

1. **Python sets rendered raw.** `tax_5_set` etc. are Python `set` objects placed
   directly into a table cell, so the PDF prints `{6307}` with braces. Join the codes
   as a comma-separated string.
2. **Doubled "only".** The tax-amount-in-words line is
   `amount_in_words(...) + ' only'`, but `amount_in_words` already ends in
   `' Paisa Only'` — the reference PDF reads `... Paisa Only only`. Emit one.
3. **Module-level mutable globals.** `generate_pdf.py` accumulates `amount`, `tax_5`,
   `total_pieces` etc. in module globals reset at the top of `generate()`. Two
   concurrent requests corrupt each other's totals. Client-side rendering plus pure
   functions removes this by construction — keep all state local.
4. **Fixed output path.** Both generators write to a single hard-coded file
   (`static/invoice_download.pdf`), so concurrent users overwrite each other and the
   "Preview" button can show someone else's invoice. Generate a Blob and hand it
   straight to the user; never persist.
5. **`'Buyer Name: ' + buyer_name` crashes on a numeric cell.** Several fields
   concatenate without coercion. Coerce every cell to string.
6. **No input validation.** A malformed sheet throws an opaque traceback. Use
   `read-excel-file`'s schema support to validate up front and show which row/column
   is wrong. Specifically warn when a GST% is not 5/12/18, since that silently
   under-taxes.
7. **Typo in the printed string** `for Orion the Corp` (§6.6, row 1). Confirm with the
   user before changing — it appears on issued invoices.

---

## 8. Suggested build order

1. **Scaffold.** `npx create-next-app@latest --typescript --app`. Confirm `npm run build`
   and a Vercel deploy work before adding anything.
2. **Extract assets** from the `legacy` branch (§2) into `public/`.
3. **Types + parsers.** `lib/types.ts`, `lib/parseInvoiceSheet.ts`,
   `lib/parsePackingSheet.ts`. Pure functions: `File → typed object`. Unit-test them
   against the sample sheets first — this is where the port is most likely to go wrong.
4. **Business logic.** `lib/calculateTax.ts`, `lib/amountInWords.ts`. Pure, unit-tested
   against the numbers in the reference PDF.
5. **Invoice PDF.** `lib/buildInvoiceDocDefinition.ts` returning a pdfmake
   document definition. Build it block by block in the §6 order, diffing against
   `/tmp/reference_invoice.pdf` as you go.
6. **Packing PDF.** `lib/buildPackingDocDefinition.ts`.
7. **UI.** Two routes (`/invoice`, `/packing`) plus a home page. Each: sample-sheet
   download link, file input, validation errors, Generate button, preview, download.
8. **Polish.** Loading states, error boundaries, mobile layout.

Keep every `lib/` function pure and free of React and pdfmake imports where possible —
that is what makes them testable.

---

## 9. Acceptance criteria

- [ ] `npm run build` passes; app deploys to Vercel with zero configuration.
- [ ] No API routes, no database, no file persistence. Verify in devtools that
      uploading a sheet issues **no network request**.
- [ ] `npm audit` reports no high/critical advisories.
- [ ] Both sample sheets from the `legacy` branch generate PDFs.
- [ ] Generated invoice matches `/tmp/reference_invoice.pdf` in **numbers** —
      subtotal, each tax slab, total, total pieces, amount in words — and closely in
      layout. Exact pixel parity is not required; correct figures are.
- [ ] Same-state `yes` produces CGST+SGST rows and the 7-column tax table;
      anything else produces IGST rows and the 5-column table. Test both.
- [ ] Packing list emits one card per box, omits zero-quantity products, and never
      splits a card across a page.
- [ ] A malformed sheet produces a readable error naming the offending row/column,
      not a stack trace.
- [ ] All seven §7 bugs are fixed (or, for #7, raised with the user).

---

## 10. Ask the user before assuming

- The company block, GSTIN, bank details, MSME number and GeM Seller ID are hard-coded
  from 2020 (§6.1). **Confirm they are still current** — the legacy source contains
  two different bank accounts (a Punjab & Sind Bank block that is built then
  discarded, and the ICICI block actually rendered). Ask which is correct.
- Whether to keep the literal `INR` or use the `₹` glyph now that the font supports it.
- Whether the `for Orion the Corp` wording is intentional.
- Whether the rigid by-position sheet format must be preserved for backward
  compatibility, or whether you may design a cleaner template. **This is the most
  consequential design decision in the port** — a schema-validated sheet would be a
  large usability win, but it invalidates every sheet the business already has.
