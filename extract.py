"""AI Invoice Extractor: turn a folder of PDF invoices into a clean Excel file.

For each PDF it:
  1. Reads the text (or, for scanned PDFs, sends the file itself to Gemini).
  2. Asks Gemini to return the invoice fields as structured JSON.
  3. Checks the math (line items vs. subtotal, subtotal vs. total).
  4. Writes everything to Excel, highlighting invoices that need review.

Usage:
    python extract.py samples                  # extract with Gemini
    python extract.py samples --dry-run        # only show the text found in each PDF
    python extract.py samples -o output/my_invoices.xlsx
"""
import argparse
import base64
import os
import sys
import time
from datetime import date
from pathlib import Path
from typing import List, Optional

import pdfplumber
from dotenv import load_dotenv
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter
from pydantic import BaseModel, Field

DEFAULT_MODEL = "gemini-3.6-flash"
FALLBACK_MODEL = "gemini-3.5-flash-lite"  # used if the main model is busy or fails
REQUEST_TIMEOUT = 90  # seconds to wait for one AI answer
MIN_TEXT_CHARS = 50   # less text than this means a scanned PDF, so send the file itself
TOLERANCE = 0.02      # allowed rounding difference when checking the math


# ---------- What we want back from the AI ----------

class LineItem(BaseModel):
    description: str = Field(description="What was sold or the service provided.")
    quantity: float = Field(description="Quantity or hours. Use 1 if not shown.")
    unit_price: float = Field(description="Price per unit as a plain number.")
    amount: float = Field(description="Line total exactly as printed.")


class Invoice(BaseModel):
    invoice_number: str = Field(description="Invoice or bill number. It may appear anywhere, including the footer.")
    vendor_name: str = Field(description="Company that issued the invoice.")
    customer_name: str = Field(description="Company or person being billed.")
    invoice_date: str = Field(description="Invoice date in YYYY-MM-DD format.")
    due_date: Optional[str] = Field(description="Due date in YYYY-MM-DD format, or null if it cannot be determined.")
    currency: str = Field(description="3-letter currency code, for example USD or EUR.")
    line_items: List[LineItem]
    subtotal: float = Field(description="Subtotal exactly as printed, before discount, tax and shipping.")
    discount: float = Field(description="Discount as a positive number, 0 if none.")
    tax: float = Field(description="Tax or VAT amount, 0 if none.")
    shipping: float = Field(description="Shipping or delivery charge, 0 if none.")
    total: float = Field(description="Final amount due exactly as printed.")


PROMPT = """Extract the data from this invoice.
Rules:
- Copy every number exactly as printed. Do NOT fix math errors; they are checked later.
- Write amounts as plain numbers (1234.5) without currency symbols or thousands separators.
- Write dates as YYYY-MM-DD. Read dates using the conventions of the invoice's country
  (for example, 05.10.2026 on a European invoice is 5 October 2026).
- If only payment terms are given (for example "Net 30"), calculate the due date from the invoice date.
- Use 0 for discount, tax or shipping when the invoice has none.
"""


# ---------- Reading and extracting ----------

def pdf_text(path: Path) -> str:
    with pdfplumber.open(path) as pdf:
        return "\n".join(page.extract_text() or "" for page in pdf.pages).strip()


def build_input(path: Path, text: str):
    if len(text) >= MIN_TEXT_CHARS:
        return f"{PROMPT}\nInvoice text:\n{text}"
    # Scanned or image-only PDF: let Gemini read the file directly.
    data = base64.b64encode(path.read_bytes()).decode("utf-8")
    return [
        {"type": "document", "data": data, "mime_type": "application/pdf"},
        {"type": "text", "text": PROMPT},
    ]


def extract_invoice(client, models: List[str], path: Path, text: str) -> Invoice:
    """Try each model twice; move to the next model if one is busy or fails."""
    request = build_input(path, text)
    last_error = None
    for model in models:
        for _ in range(2):
            try:
                interaction = client.interactions.create(
                    model=model,
                    input=request,
                    response_format={
                        "type": "text",
                        "mime_type": "application/json",
                        "schema": Invoice.model_json_schema(),
                    },
                    timeout=REQUEST_TIMEOUT,
                )
                return Invoice.model_validate_json(interaction.output_text)
            except Exception as err:  # busy model, timeout or unusable answer
                last_error = err
                print(f"   {model} failed ({type(err).__name__}), trying again...")
                time.sleep(5)
    raise last_error


# ---------- Checking the math ----------

def check_invoice(inv: Invoice) -> str:
    problems = []
    for item in inv.line_items:
        if abs(item.quantity * item.unit_price - item.amount) > TOLERANCE:
            problems.append(f"'{item.description}': {item.quantity:g} x {item.unit_price:,.2f} is not {item.amount:,.2f}")

    items_sum = round(sum(item.amount for item in inv.line_items), 2)
    if abs(items_sum - inv.subtotal) > TOLERANCE:
        problems.append(f"line items add up to {items_sum:,.2f} but subtotal says {inv.subtotal:,.2f}")

    expected_total = round(inv.subtotal - inv.discount + inv.tax + inv.shipping, 2)
    if abs(expected_total - inv.total) > TOLERANCE:
        problems.append(f"subtotal - discount + tax + shipping = {expected_total:,.2f} but total says {inv.total:,.2f}")

    return "OK" if not problems else "Needs review: " + "; ".join(problems)


# ---------- Writing Excel ----------

INVOICE_COLUMNS = [
    ("File", 28), ("Invoice #", 15), ("Vendor", 28), ("Customer", 26),
    ("Invoice Date", 13), ("Due Date", 13), ("Currency", 9),
    ("Subtotal", 12), ("Discount", 11), ("Tax", 11), ("Shipping", 11), ("Total", 12),
    ("Check", 70),
]
ITEM_COLUMNS = [
    ("File", 28), ("Invoice #", 15), ("Description", 38),
    ("Quantity", 10), ("Unit Price", 12), ("Amount", 12),
]
MONEY_FORMAT = "#,##0.00"
DATE_FORMAT = "yyyy-mm-dd"
REVIEW_FILL = PatternFill("solid", start_color="FFF4CCCC")


def setup_sheet(ws, columns):
    ws.append([name for name, _ in columns])
    for i, (_, width) in enumerate(columns, start=1):
        ws.column_dimensions[get_column_letter(i)].width = width
        ws.cell(row=1, column=i).font = Font(bold=True)
    ws.freeze_panes = "A2"


def as_date(value: Optional[str]):
    try:
        return date.fromisoformat(value) if value else None
    except ValueError:
        return value  # keep the text if it isn't a clean date


def save_excel(results, out_path: Path):
    wb = Workbook()
    inv_ws = wb.active
    inv_ws.title = "Invoices"
    items_ws = wb.create_sheet("Line Items")
    setup_sheet(inv_ws, INVOICE_COLUMNS)
    setup_sheet(items_ws, ITEM_COLUMNS)

    for file_name, inv, status in results:
        if inv is None:
            inv_ws.append([file_name] + [None] * 11 + [status])
        else:
            inv_ws.append([
                file_name, inv.invoice_number, inv.vendor_name, inv.customer_name,
                as_date(inv.invoice_date), as_date(inv.due_date), inv.currency,
                inv.subtotal, inv.discount, inv.tax, inv.shipping, inv.total, status,
            ])
            for item in inv.line_items:
                items_ws.append([file_name, inv.invoice_number, item.description,
                                 item.quantity, item.unit_price, item.amount])
        row = inv_ws[inv_ws.max_row]
        row[4].number_format = row[5].number_format = DATE_FORMAT
        for cell in row[7:12]:
            cell.number_format = MONEY_FORMAT
        if status != "OK":
            for cell in row:
                cell.fill = REVIEW_FILL

    for row in items_ws.iter_rows(min_row=2, min_col=5, max_col=6):
        for cell in row:
            cell.number_format = MONEY_FORMAT

    out_path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(out_path)


# ---------- Main ----------

def main():
    sys.stdout.reconfigure(encoding="utf-8", errors="replace", line_buffering=True)
    parser = argparse.ArgumentParser(description="Extract invoice data from PDFs into Excel using Gemini.")
    parser.add_argument("folder", type=Path, help="folder with PDF invoices")
    parser.add_argument("-o", "--output", type=Path, default=Path("output/invoices.xlsx"), help="Excel file to create")
    parser.add_argument("--dry-run", action="store_true", help="only show the text found in each PDF (no AI call)")
    args = parser.parse_args()

    pdfs = sorted(args.folder.glob("*.pdf"))
    if not pdfs:
        sys.exit(f"No PDF files found in {args.folder}")

    if args.dry_run:
        for path in pdfs:
            text = pdf_text(path)
            mode = "text found" if len(text) >= MIN_TEXT_CHARS else "no text: will send the PDF itself"
            print(f"\n=== {path.name} ({len(text)} characters, {mode}) ===")
            print(text[:500])
        return

    load_dotenv()
    if not os.getenv("GEMINI_API_KEY"):
        sys.exit("GEMINI_API_KEY is missing. Copy .env.example to .env and paste your key into it.")

    from google import genai  # imported here so --dry-run works without the SDK set up
    client = genai.Client()
    main_model = os.getenv("GEMINI_MODEL") or DEFAULT_MODEL
    models = [main_model] if main_model == FALLBACK_MODEL else [main_model, FALLBACK_MODEL]

    results = []
    for path in pdfs:
        print(f"Reading {path.name} ...")
        try:
            inv = extract_invoice(client, models, path, pdf_text(path))
            status = check_invoice(inv)
        except Exception as err:
            inv, status = None, f"Error: {str(err)[:200]}"
        print(f"   {status}")
        results.append((path.name, inv, status))

    save_excel(results, args.output)
    ok = sum(1 for _, _, status in results if status == "OK")
    print(f"\nDone: {len(results)} invoices, {ok} OK, {len(results) - ok} need review.")
    print(f"Saved to {args.output}")


if __name__ == "__main__":
    main()
