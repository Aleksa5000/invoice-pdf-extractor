# invoice-pdf-extractor

[![CI](https://github.com/Aleksa5000/invoice-pdf-extractor/actions/workflows/ci.yml/badge.svg)](https://github.com/Aleksa5000/invoice-pdf-extractor/actions/workflows/ci.yml)

**Turn a folder of PDF invoices into a clean, validated Excel sheet in seconds - no more retyping.**

## The problem it solves

Accounts-payable and bookkeeping teams still copy invoice numbers, dates, vendors and totals from PDFs into spreadsheets by hand. It is slow, boring, and a single mistyped digit ends up in your books.

This tool reads every PDF in a folder, pulls out the header fields and the line-item table, and **checks the maths for you**: do the line items add up to the subtotal, and does subtotal + tax equal the total? Invoices that pass are marked `OK`; the rest are marked `NEEDS REVIEW` with the exact reason, so a human only looks at the ones that need it.

## Features

- Batch-processes a whole folder of text-based PDF invoices ([pdfplumber](https://github.com/jsvine/pdfplumber))
- Extracts: invoice number, date, due date, vendor, currency, subtotal, tax, total, and line items (description, qty, unit price, amount)
- Validates line-items sum vs. subtotal and subtotal + tax vs. total (small rounding tolerance)
- Flags each invoice `OK` / `NEEDS REVIEW` with a human-readable reason (missing fields, unreadable PDFs, mismatches)
- Outputs Excel (`Invoices` + `Line items` sheets) and optionally JSON
- A corrupt PDF never stops the batch - it is just flagged

## Example

```
$ python extract.py sample_invoices -o output.xlsx --json output.json
acme_supplies.pdf          INV-2026-0101       415.92  OK
bluecloud_hosting.pdf      BC-77412            170.46  OK
greenleaf_catering.pdf     GL-0589           1,273.47  OK
mismatch_total.pdf         KL-30045          1,510.15  NEEDS REVIEW  (subtotal + tax 1410.15 != total 1510.15)
nordic_design_studio.pdf   2026/033          2,275.00  OK

5 invoices -> output.xlsx
JSON -> output.json
4 OK, 1 need review
```

Extracted rows (sheet `Invoices`):

| Invoice # | Vendor | Date | Due date | Cur. | Subtotal | Tax | Total | Status |
|---|---|---|---|---|---:|---:|---:|---|
| INV-2026-0101 | Acme Office Supplies Ltd | 2026-03-04 | 2026-04-03 | GBP | 346.60 | 69.32 | 415.92 | OK |
| BC-77412 | BlueCloud Hosting Inc. | 2026-03-10 | 2026-03-24 | USD | 157.47 | 12.99 | 170.46 | OK |
| GL-0589 | Greenleaf Catering Co. | 2026-03-21 | 2026-04-20 | EUR | 1,122.00 | 151.47 | 1,273.47 | OK |
| KL-30045 | Kestrel Logistics GmbH | 2026-03-28 | 2026-04-27 | EUR | 1,185.00 | 225.15 | 1,510.15 | **NEEDS REVIEW** |
| 2026/033 | Nordic Design Studio AB | 2026-03-15 | 2026-04-14 | EUR | 1,820.00 | 455.00 | 2,275.00 | OK |

The Kestrel invoice is deliberately wrong (the printed total is 100.00 too high) to show the validation catching it.

## Quickstart

```bash
python -m pip install -r requirements.txt
python extract.py sample_invoices/ -o output.xlsx            # Excel
python extract.py sample_invoices/ -o output.xlsx --json out.json   # Excel + JSON
```

The `sample_invoices/` folder is included; regenerate it with `python make_samples.py`. Point the script at your own folder of PDFs to use it on real invoices.

## Running tests

```bash
python -m pip install -r requirements.txt pytest
python -m pytest -q
```

Tests generate invoices into a temp folder, run the extraction, and assert the fields and that the mismatch invoice is flagged. No network needed.

## Limitations

- **Text-based PDFs only.** Scanned invoices (images) need OCR first, which is out of scope here.
- Field extraction uses simple regexes tuned to a common layout ("Invoice No:", "Due Date:", "Subtotal:", ...) plus the first table on the page. Real vendors differ, so for a client project the patterns are adapted to your actual invoice formats.
- Dates are expected as `YYYY-MM-DD`; amounts as `1,234.56`.

## Tech stack

Python 3.11+, pdfplumber, openpyxl, fpdf2 (sample generation), pytest, GitHub Actions.

## Need something similar?

Need invoices, receipts, or other documents extracted into your spreadsheet or system? **Contact me on Upwork** - Aleksa Djorovic, Python automation.

## License

MIT
