"""Batch-extract structured data from text-based PDF invoices into Excel / JSON."""
import argparse
import json
import re
from pathlib import Path

import pdfplumber
from openpyxl import Workbook

TOLERANCE = 0.02  # allowed rounding difference
DATE = r"(\d{4}-\d{2}-\d{2})"
NUM = r"-?[\d,]+\.\d{2}"


def num(s):
    return float(s.replace(",", ""))


def find(pattern, text):
    m = re.search(pattern, text, re.I | re.M)
    return m.group(1) if m else None


def extract_invoice(path):
    with pdfplumber.open(path) as pdf:
        text = "\n".join(p.extract_text() or "" for p in pdf.pages)
        tables = [t for p in pdf.pages for t in p.extract_tables()]

    # Line items: first table, skip the header row, keep rows with numeric cells
    items = []
    for row in (tables[0][1:] if tables else []):
        try:
            items.append({"description": row[0], "qty": num(row[1]), "unit_price": num(row[2]),
                          "amount": num(row[3])})
        except (ValueError, IndexError, AttributeError):
            continue

    def money(label):
        v = find(rf"^{label}[^:\n]*:\s*({NUM})", text)
        return num(v) if v else None

    inv = {
        "file": Path(path).name,
        "vendor": next((l.strip() for l in text.splitlines() if l.strip()), None),  # first line
        "invoice_number": find(r"Invoice (?:No|Number|#)\.?:?\s*(\S+)", text),
        "date": find(rf"^Date:\s*{DATE}", text),
        "due_date": find(rf"Due Date:\s*{DATE}", text),
        "currency": find(r"Currency:\s*([A-Z]{3})", text),
        "subtotal": money("Subtotal"),
        "tax": money("Tax"),
        "total": money("Total"),
        "line_items": items,
    }
    inv["status"], inv["reason"] = validate(inv)
    return inv


def validate(inv):
    problems = [f"missing {k}" for k in ("invoice_number", "date", "vendor", "subtotal", "tax", "total")
                if inv[k] is None]
    if not inv["line_items"]:
        problems.append("no line items found")
    if not problems:
        items_sum = round(sum(i["amount"] for i in inv["line_items"]), 2)
        if abs(items_sum - inv["subtotal"]) > TOLERANCE:
            problems.append(f"line items sum {items_sum:.2f} != subtotal {inv['subtotal']:.2f}")
        if abs(inv["subtotal"] + inv["tax"] - inv["total"]) > TOLERANCE:
            problems.append(f"subtotal + tax {inv['subtotal'] + inv['tax']:.2f} != total {inv['total']:.2f}")
    return ("NEEDS REVIEW", "; ".join(problems)) if problems else ("OK", "")


def extract_folder(folder):
    invoices = []
    for pdf in sorted(Path(folder).glob("*.pdf")):
        try:
            invoices.append(extract_invoice(pdf))
        except Exception as e:  # keep the batch going on a corrupt file
            invoices.append({"file": pdf.name, "line_items": [], "status": "NEEDS REVIEW",
                             "reason": f"could not read PDF: {e}"})
    return invoices


COLS = ["file", "invoice_number", "vendor", "date", "due_date", "currency", "subtotal", "tax", "total",
        "status", "reason"]
ITEM_COLS = ["description", "qty", "unit_price", "amount"]


def write_xlsx(invoices, path):
    wb = Workbook()
    ws = wb.active
    ws.title = "Invoices"
    ws.append(COLS)
    items = wb.create_sheet("Line items")
    items.append(["file", "invoice_number"] + ITEM_COLS)
    for inv in invoices:
        ws.append([inv.get(c) for c in COLS])
        for it in inv["line_items"]:
            items.append([inv["file"], inv.get("invoice_number")] + [it[c] for c in ITEM_COLS])
    for sheet in (ws, items):
        sheet.freeze_panes = "A2"
        for col in sheet.columns:
            sheet.column_dimensions[col[0].column_letter].width = min(
                max(len(str(c.value or "")) for c in col) + 2, 50)
    wb.save(path)


def main():
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("folder", help="folder with PDF invoices")
    ap.add_argument("-o", "--output", default="invoices.xlsx", help="Excel output path")
    ap.add_argument("--json", help="also write JSON to this path")
    args = ap.parse_args()

    invoices = extract_folder(args.folder)
    for inv in invoices:
        total = inv.get("total")
        print(f"{inv['file']:<26} {inv.get('invoice_number') or '-':<15} "
              f"{'-' if total is None else format(total, ',.2f'):>10}  {inv['status']}"
              + (f"  ({inv['reason']})" if inv["reason"] else ""))
    write_xlsx(invoices, args.output)
    print(f"\n{len(invoices)} invoices -> {args.output}")
    if args.json:
        Path(args.json).write_text(json.dumps(invoices, indent=2), encoding="utf-8")
        print(f"JSON -> {args.json}")
    ok = sum(i["status"] == "OK" for i in invoices)
    print(f"{ok} OK, {len(invoices) - ok} need review")


if __name__ == "__main__":
    main()
