# ORION_CORP_INVOICE_GENERATOR

A Django web app that turns an Excel spreadsheet into a formatted PDF. It has two
generators:

- **Invoice Generator** — builds a tax invoice PDF (GST slabs of 5% / 12% / 18%,
  totals, amount in words) from a bill spreadsheet.
- **Packing List** — builds a packing list PDF from a packing spreadsheet.

PDFs are rendered with `reportlab`; spreadsheets are read with `pandas` + `openpyxl`.

---

## Running it locally

### Prerequisites

- Python 3.12 (3.10+ should also work)
- macOS / Linux instructions below; on Windows swap the activate command

### 1. Clone and enter the Django project directory

```bash
git clone <repo-url>
cd ORION_CORP_INVOICE_GENERATOR/orion_website
```

All commands below run from `orion_website/` — the directory containing `manage.py`.

### 2. Create and activate a virtual environment

```bash
python3.12 -m venv venv
source venv/bin/activate
```

On Windows use `venv\Scripts\activate` instead.

To leave the environment later, run `deactivate`.

### 3. Install dependencies

```bash
pip install --upgrade pip
pip install -r requirements.txt
```

### 4. Apply database migrations

```bash
python manage.py migrate
```

This sets up the bundled SQLite database (`db.sqlite3`). The app itself defines no
models yet — the migrations only cover Django's built-in admin/auth/session tables.

### 5. Run the development server

```bash
python manage.py runserver
```

Open <http://127.0.0.1:8000/> (or <http://localhost:8000/>). Stop the server with
`Ctrl+C`.

### 6. (Optional) Create an admin user

```bash
python manage.py createsuperuser
```

Then log in at <http://127.0.0.1:8000/admin/>.

---

## Using the app

The home page links to both generators.

| Page | URL |
| --- | --- |
| Home | `/` |
| Invoice Generator | `/invoice_generator/invoice_upload` |
| Packing List | `/packing_list/packing_upload` |
| Django admin | `/admin/` |

On either generator page:

1. Click **Sample Sheet** to download the template spreadsheet
   (`static/invoice_sample.xlsx` or `static/packing_sample.xlsx`).
2. Fill in your rows, keeping the header row and column order intact.
3. Choose the file and click **Upload**.
4. Click **Preview** to open the freshly generated PDF.

The generated PDFs are written back into `static/`:

- `static/invoice_download.pdf`
- `static/packing_download.pdf`

Each upload **overwrites** the previous PDF, so download a result before generating
the next one.

### Invoice spreadsheet layout

The invoice sheet is read strictly by position, so keep every row where it is and
only change the values. Start from `static/invoice_sample.xlsx` rather than building
a sheet from scratch.

Counting the spreadsheet's own header as row 1:

| Sheet row | Contents |
| --- | --- |
| 1 | Labels: invoice type, invoice number, buyer name, buyer address, buyer phone, buyer GST, bill date, delivery note |
| 2 | **Values** for the labels in row 1 |
| 3 | Labels: mode of payment, supplier ref, other ref, buyer order number, buyer order date, despatch document number, delivery note date, despatched through |
| 4 | **Values** for the labels in row 3 |
| 5 | Labels: destination, term1, term2, term3, same state (`yes`/`no`) |
| 6 | **Values** for the labels in row 5 |
| 7 | Labels for the line-item table |
| 8 onward | **One line item per row** |

Each line-item row holds eight columns, in this order:

| Column | Meaning | Notes |
| --- | --- | --- |
| A | SI | Serial number |
| B | Description of goods | |
| C | HSN/SAC | Grouped per tax slab in the tax summary |
| D | Quantity | Summed into total pieces |
| E | Rate | |
| F | per | Unit, e.g. `pcs` |
| G | GST % | Must be `5`, `12`, or `18` |
| H | Amount | Summed into the taxable value |

Two details worth knowing:

- A row whose GST % is not exactly 5, 12, or 18 still counts toward the amount total
  but contributes no tax.
- The **same state** cell in row 6 decides the tax split: `yes` prints CGST + SGST at
  half the rate each, anything else prints a single IGST line.

The packing list sheet follows the same "labels then values, then item rows" idea —
again, start from `static/packing_sample.xlsx`.

---

## Project layout

```
orion_website/
├── manage.py
├── requirements.txt
├── db.sqlite3
├── orion_website/            # Django project config (settings, root urls, wsgi)
├── invoice_generator/        # Invoice app
│   ├── views.py              # Reads the upload into a DataFrame
│   └── generate_pdf.py       # Builds the invoice PDF with reportlab
├── packing_list/             # Packing list app
│   ├── views.py
│   └── generate_packing_list.py
├── templates/                # HTML templates for both apps
├── static/                   # Sample sheets, signature image, generated PDFs
└── media/
```

---

## Notes and known limitations

- `DEBUG = True` and the `SECRET_KEY` is committed in `orion_website/settings.py`.
  That is fine for local development, but both must change before any real
  deployment.
- There is no login yet — anyone who can reach the site can generate a PDF.
- Generated PDFs are written to a fixed path in `static/`, so concurrent uploads
  would overwrite each other. `generate_pdf.py` also accumulates its totals in
  module-level globals (reset at the start of each run), which is not safe under
  concurrent requests.
- Uploaded spreadsheets are parsed in memory and never saved.

---

## Roadmap

The original staged plan for this project:

- **Stage 1** — generate a proforma invoice PDF locally from hard-coded values. ✅
- **Stage 2** — generate proforma and tax invoices from an Excel file, extracting
  the data with pandas. ✅
- **Stage 3** — expand to all document formats. *(packing list added)*
- **Stage 4** — a secure Django form with proper user login, deployed to a public
  server, offering two paths: fill the form manually, or upload a pre-templated
  Excel sheet and fill in only the essentials. *(the Django app and Excel upload
  exist; login and hardening do not)*
- **Stage 5** — full support for the above, plus logging records for later
  tallying.
