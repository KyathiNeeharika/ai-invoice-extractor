"""Create 5 fake invoice PDFs for testing the extractor.

Every company, person and number here is made up. The invoices use different
layouts, date formats and currencies on purpose, and one has a math error,
so the extractor has something realistic to deal with.

Usage:
    python make_samples.py
"""
from pathlib import Path

from reportlab.lib.pagesizes import A4, letter
from reportlab.pdfgen import canvas

OUT_DIR = Path(__file__).parent / "samples"


def usd(x):
    return f"-${abs(x):,.2f}" if x < 0 else f"${x:,.2f}"


def eur(x):
    # European style: 1.234,50 €
    s = f"{x:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    return f"{s} €"


def qty_text(q):
    return str(int(q)) if float(q).is_integer() else str(q)


INVOICES = [
    {
        "file": "brightpath_BP-10234.pdf",
        "page": letter,
        "money": usd,
        "vendor": ["BrightPath Office Supplies", "418 Elm Street, Springfield, IL 62701", "(555) 201-3344"],
        "title": "INVOICE",
        "meta": [("Invoice #:", "BP-10234"), ("Date:", "2026-10-01"), ("Due:", "2026-10-31")],
        "bill_to": ["Oakridge Dental Clinic", "92 Maple Ave, Springfield, IL 62704"],
        "items": [
            ("Printer paper, A4, 500 sheets", 10, 6.50),
            ("Black ink cartridge XL", 4, 32.00),
            ("Ergonomic office chair", 2, 189.00),
            ("Desk organizer set", 3, 14.75),
        ],
        "totals": [("Subtotal", 615.25), ("Tax (8%)", 49.22), ("Shipping", 25.00), ("TOTAL", 689.47)],
        "footer": ["Thank you for your business!"],
    },
    {
        "file": "cedarlane_CL-2026-0457.pdf",
        "page": letter,
        "money": usd,
        "vendor": ["Cedar Lane Catering Co.", "7 Harbor Road, Portland, OR 97205", "orders@cedarlane.example"],
        "title": "TAX INVOICE",
        "meta": [("Bill No.", "CL/2026/0457"), ("Issued:", "Oct 3, 2026"), ("Due Date:", "18 Oct 2026")],
        "bill_to": ["Summit Realty Group", "Attn: Dana Brooks", "1500 Pine St, Portland, OR 97201"],
        "items": [
            ("Lunch buffet (per person)", 45, 18.00),
            ("Coffee & tea service (per person)", 45, 3.50),
            ("Dessert platter", 3, 42.00),
            ("Event staff (hours)", 6, 28.00),
        ],
        "totals": [("Subtotal", 1261.50), ("Sales tax 7%", 88.31), ("Amount Due", 1349.81)],
        "footer": ["Please include the bill number with your payment."],
    },
    {
        "file": "lumen_LWS-0088.pdf",
        "page": letter,
        "money": usd,
        "vendor": ["Lumen Web Studio", "33 Canal St, Austin, TX 78701"],
        "title": "INVOICE",
        "meta": [("Inv. No:", "LWS-0088"), ("Date:", "10/05/2026"), ("Payment terms:", "Net 30")],
        "bill_to": ["Greenleaf Yoga Studio", "210 Oak Blvd, Austin, TX 78704"],
        "items": [
            ("Website redesign (hours)", 24, 65.00),
            ("Booking page integration (hours)", 8, 65.00),
            ("Hosting setup (one-time)", 1, 150.00),
        ],
        "totals": [("Subtotal", 2230.00), ("Discount (10%)", -223.00), ("Total due", 2007.00)],
        "footer": ["No sales tax charged on these services."],
    },
    {
        "file": "bluegate_eur_invoice.pdf",
        "page": A4,
        "money": eur,
        "vendor": ["Bluegate IT Services Ltd", "Kanalstraat 14, 1012 Amsterdam, Netherlands"],
        "title": "INVOICE",
        "meta": [("Date:", "12.10.2026"), ("Payable by:", "26.10.2026")],
        "bill_to": ["Kestrel Design Studio Ltd", "Rue des Fleurs 8, 1000 Brussels, Belgium"],
        "items": [
            ("Managed IT support (monthly)", 1, 450.00),
            ("Laptop setup & configuration", 5, 60.00),
            ("Cloud email licenses", 12, 11.50),
        ],
        "totals": [("Subtotal", 888.00), ("VAT 20%", 177.60), ("Total", 1065.60)],
        # The invoice number is only in the footer, to test unusual layouts.
        "footer": ["Invoice Number: BG-2026-1147", "Bank: NL00 EXMP 0123 4567 89 (example only)"],
    },
    {
        "file": "riverbend_RPS-5521.pdf",
        "page": letter,
        "money": usd,
        "vendor": ["Riverbend Print Shop", "5 Mill Lane, Burlington, VT 05401"],
        "title": "INVOICE",
        "meta": [("Invoice No:", "RPS-5521"), ("Date:", "November 2, 2026"), ("Due:", "November 16, 2026")],
        "bill_to": ["Bayside Bakery", "77 Shore Rd, Burlington, VT 05401"],
        "items": [
            ("Business cards (box of 500)", 4, 45.00),
            ("Flyers A5, full colour (1000)", 2, 120.00),
            ("Roll-up banner", 1, 220.00),
        ],
        # Deliberate mistake: the items add up to 640.00, not 680.00.
        "totals": [("Subtotal", 680.00), ("Tax 6%", 40.80), ("TOTAL", 720.80)],
        "footer": ["Thank you!"],
    },
]


def draw_invoice(spec):
    path = OUT_DIR / spec["file"]
    width, height = spec["page"]
    money = spec["money"]
    c = canvas.Canvas(str(path), pagesize=spec["page"])
    left, right = 50, width - 50

    # Vendor block (top left)
    c.setFont("Helvetica-Bold", 16)
    c.drawString(left, height - 60, spec["vendor"][0])
    c.setFont("Helvetica", 9)
    for i, line in enumerate(spec["vendor"][1:]):
        c.drawString(left, height - 76 - i * 12, line)

    # Title and invoice details (top right)
    c.setFont("Helvetica-Bold", 20)
    c.drawRightString(right, height - 60, spec["title"])
    c.setFont("Helvetica", 10)
    for i, (label, value) in enumerate(spec["meta"]):
        y = height - 84 - i * 14
        c.drawRightString(right - 110, y, label)
        c.drawRightString(right, y, value)

    # Bill to
    y = height - 160
    c.setFont("Helvetica-Bold", 10)
    c.drawString(left, y, "Bill To:")
    c.setFont("Helvetica", 10)
    for i, line in enumerate(spec["bill_to"]):
        c.drawString(left, y - 14 - i * 13, line)

    # Line items table
    y = height - 250
    c.setFont("Helvetica-Bold", 10)
    c.drawString(left, y, "Description")
    c.drawRightString(right - 190, y, "Qty")
    c.drawRightString(right - 100, y, "Unit Price")
    c.drawRightString(right, y, "Amount")
    c.line(left, y - 5, right, y - 5)
    c.setFont("Helvetica", 10)
    for desc, qty, price in spec["items"]:
        y -= 18
        c.drawString(left, y, desc)
        c.drawRightString(right - 190, y, qty_text(qty))
        c.drawRightString(right - 100, y, money(price))
        c.drawRightString(right, y, money(round(qty * price, 2)))
    c.line(left, y - 8, right, y - 8)

    # Totals (right side)
    y -= 28
    for label, value in spec["totals"]:
        is_final = label == spec["totals"][-1][0]
        c.setFont("Helvetica-Bold" if is_final else "Helvetica", 11 if is_final else 10)
        c.drawRightString(right - 110, y, label)
        c.drawRightString(right, y, money(value))
        y -= 16

    # Footer
    c.setFont("Helvetica", 8)
    for i, line in enumerate(spec["footer"]):
        c.drawString(left, 60 - i * 11, line)

    c.save()
    return path


def main():
    OUT_DIR.mkdir(exist_ok=True)
    for spec in INVOICES:
        print("Created", draw_invoice(spec).name)


if __name__ == "__main__":
    main()
