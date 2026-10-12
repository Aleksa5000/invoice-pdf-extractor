"""Generate realistic sample invoices (PDF) into sample_invoices/."""
import sys
from pathlib import Path

from fpdf import FPDF

# (file, vendor, address, number, date, due, currency, tax_rate, items, extra_added_to_total)
SAMPLES = [
    ("acme_supplies.pdf", "Acme Office Supplies Ltd", "12 Market Street, London", "INV-2026-0101",
     "2026-03-04", "2026-04-03", "GBP", 0.20,
     [("A4 Copy Paper (box of 5 reams)", 10, 24.50), ("Ballpoint pens (pack of 50)", 4, 11.90),
      ("Desk organiser", 3, 18.00)], None),
    ("bluecloud_hosting.pdf", "BlueCloud Hosting Inc.", "800 Cloud Ave, Austin, TX", "BC-77412",
     "2026-03-10", "2026-03-24", "USD", 0.0825,
     [("VPS Standard - March 2026", 3, 40.00), ("Managed backups", 3, 9.99),
      ("Dedicated IPv4 address", 3, 2.50)], None),
    ("nordic_design_studio.pdf", "Nordic Design Studio AB", "Storgatan 5, Stockholm", "2026/033",
     "2026-03-15", "2026-04-14", "EUR", 0.25,
     [("Brand identity workshop (hours)", 8, 95.00), ("Logo design - 3 concepts", 1, 640.00),
      ("Social media templates", 12, 35.00)], None),
    ("greenleaf_catering.pdf", "Greenleaf Catering Co.", "47 Orchard Road, Dublin", "GL-0589",
     "2026-03-21", "2026-04-20", "EUR", 0.135,
     [("Lunch buffet (per person)", 40, 18.50), ("Coffee & pastries (per person)", 40, 6.25),
      ("Service staff (hours)", 6, 22.00)], None),
    # Deliberate mismatch: the printed total is 100.00 too high
    ("mismatch_total.pdf", "Kestrel Logistics GmbH", "Hafenstrasse 9, Hamburg", "KL-30045",
     "2026-03-28", "2026-04-27", "EUR", 0.19,
     [("Pallet transport Hamburg-Berlin", 6, 185.00), ("Fuel surcharge", 6, 12.50)], 100.00),
]


def money(x):
    return f"{x:,.2f}"


def make_invoice(path, vendor, address, number, date, due, currency, rate, items, extra=None):
    subtotal = round(sum(q * p for _, q, p in items), 2)
    tax = round(subtotal * rate, 2)
    total = round(subtotal + tax + (extra or 0), 2)
    pdf = FPDF()
    pdf.add_page()
    pdf.set_font("Helvetica", "B", 16)
    pdf.cell(0, 8, vendor, new_x="LMARGIN", new_y="NEXT")
    pdf.set_font("Helvetica", size=10)
    pdf.cell(0, 6, address, new_x="LMARGIN", new_y="NEXT")
    pdf.ln(6)
    pdf.set_font("Helvetica", "B", 14)
    pdf.cell(0, 8, "INVOICE", new_x="LMARGIN", new_y="NEXT")
    pdf.set_font("Helvetica", size=10)
    for line in (f"Invoice No: {number}", f"Date: {date}", f"Due Date: {due}", f"Currency: {currency}"):
        pdf.cell(0, 6, line, new_x="LMARGIN", new_y="NEXT")
    pdf.ln(6)
    with pdf.table(col_widths=(90, 15, 30, 30), text_align=("LEFT", "RIGHT", "RIGHT", "RIGHT")) as t:
        r = t.row()
        for h in ("Description", "Qty", "Unit Price", "Amount"):
            r.cell(h)
        for desc, q, p in items:
            r = t.row()
            for v in (desc, str(q), money(p), money(q * p)):
                r.cell(v)
    pdf.ln(6)
    for label, val in (("Subtotal", subtotal), (f"Tax ({rate * 100:g}%)", tax), ("Total", total)):
        pdf.cell(0, 6, f"{label}: {money(val)}", align="R", new_x="LMARGIN", new_y="NEXT")
    pdf.ln(8)
    pdf.cell(0, 6, "Payment by bank transfer within the due date. Thank you for your business.")
    pdf.output(str(path))


def main(out="sample_invoices"):
    out = Path(out)
    out.mkdir(exist_ok=True)
    for name, *args in SAMPLES:
        make_invoice(out / name, *args)
        print("created", out / name)


if __name__ == "__main__":
    main(*sys.argv[1:])
