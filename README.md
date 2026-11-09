# AI Invoice Extractor

Turn a folder of PDF invoices into a clean Excel spreadsheet in seconds, using Google Gemini. It checks the math on every invoice and highlights the ones that need a human to look at them.

<!-- Demo video: add your Loom/YouTube link here -->

## Why

Small businesses and bookkeepers still type invoice details into spreadsheets by hand. Rule-based tools break whenever a vendor changes their layout. This tool reads **any layout** with AI, then double-checks the numbers so mistakes don't slip through.

## What it does

- **Reads any invoice layout.** No templates or rules to maintain. It finds the invoice number even when it's tucked in the footer.
- **Handles different formats.** Dates like `2026-10-01`, `Oct 3, 2026`, `10/05/2026` and `12.10.2026`; currencies like `$1,261.50` and `1.065,60 €`.
- **Works out due dates** from payment terms such as "Net 30".
- **Catches errors.** It checks that line items add up to the subtotal, and that subtotal − discount + tax + shipping equals the total. Invoices that don't add up are highlighted in red.
- **Works on scanned PDFs.** If a PDF has no text layer, the file itself is sent to Gemini to read.
- **Keeps going when the AI is busy.** If the main model is overloaded, it retries and falls back to a backup model.
- **Exports to Excel** with two sheets: `Invoices` (one row per invoice) and `Line Items` (one row per item).

## Example results

From the 5 sample invoices in [`samples/`](samples), all of which are fake:

| Invoice # | Vendor | Date | Due | Currency | Total | Check |
|---|---|---|---|---|---:|---|
| BG-2026-1147 | Bluegate IT Services Ltd | 2026-10-12 | 2026-10-26 | EUR | 1,065.60 | OK |
| BP-10234 | BrightPath Office Supplies | 2026-10-01 | 2026-10-31 | USD | 689.47 | OK |
| CL/2026/0457 | Cedar Lane Catering Co. | 2026-10-03 | 2026-10-18 | USD | 1,349.81 | OK |
| LWS-0088 | Lumen Web Studio | 2026-10-05 | 2026-11-04 *(Net 30)* | USD | 2,007.00 | OK |
| RPS-5521 | Riverbend Print Shop | 2026-11-02 | 2026-11-16 | USD | 720.80 | **Needs review:** line items add up to 640.00 but subtotal says 680.00 |

The full output is in [`output/invoices.xlsx`](output/invoices.xlsx).

## How it works

```
PDF ──► pdfplumber (read text) ──► Gemini (structured JSON, Pydantic schema) ──► math checks ──► Excel (openpyxl)
          │ no text found?
          └──► send the PDF itself to Gemini
```

The AI's answer is forced into a fixed schema (invoice number, dates, currency, line items, subtotal, discount, tax, shipping, total). It's told to copy numbers exactly as printed, so the math checks can catch real errors instead of the AI quietly fixing them.

## Quick start

Requires Python 3.10+.

```bash
pip install -r requirements.txt
```

Get a free Gemini API key at [aistudio.google.com](https://aistudio.google.com). Copy `.env.example` to `.env` and paste the key after `GEMINI_API_KEY=`.

```bash
python make_samples.py      # creates 5 fake test invoices in samples/
python extract.py samples   # writes output/invoices.xlsx
```

### Options

```bash
python extract.py my_invoices -o output/october.xlsx   # choose the output file
python extract.py samples --dry-run                    # just show the text found in each PDF, no AI call
```

## Project structure

```
extract.py        main script: PDFs -> AI -> checks -> Excel
make_samples.py   generates the fake sample invoices
samples/          5 sample invoices (made up)
output/           example Excel output
.env.example      settings template (copy to .env)
```

## Notes

- All sample companies, people and numbers are fictional.
- Gemini's **free tier** may use your inputs to improve Google's products. For real client documents, use a paid API key, which isn't used for training, or swap in another AI provider.
- Built with Python, [google-genai](https://pypi.org/project/google-genai/), [pdfplumber](https://github.com/jsvine/pdfplumber), [Pydantic](https://docs.pydantic.dev/) and [openpyxl](https://openpyxl.readthedocs.io/).

## About

Built by **Kyathi**, Data Analyst & AI Automation.
Need invoices, receipts or other documents turned into clean data? Contact me on [Upwork](https://www.upwork.com/freelancers/~0169cd5388c4a1e327) or [LinkedIn](https://www.linkedin.com/in/kyathi-neeharika).
