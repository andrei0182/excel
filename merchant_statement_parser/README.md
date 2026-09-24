# Fiserv / CardPointe merchant statement parser

Turns a Fiserv "Your Card Processing Statement" PDF into clean tables and reconciles every section
against the others, so parsing errors (and odd statement data) show up automatically.

```bash
pip install -r requirements.txt
python parse_fiserv_statement.py statement.pdf                      # -> statement_parsed.xlsx
python parse_fiserv_statement.py statement.pdf -o out.xlsx --csv out_csv --json out.json
python parse_fiserv_statement.py folder_of_pdfs/ -o parsed/         # batch
```

Exit code is 1 if any check is MISMATCH or MISSING, so the script can gate an automated pipeline.

## Output (Excel sheets / CSV files)

| Sheet | Content |
|---|---|
| Statement_Info | period, corporate number, locations, merchant block, account summary, customer KPIs, YoY / MoM |
| Validation | ~80 reconciliation checks: OK / MISMATCH / MISSING / REVIEW / INFO |
| Daily_Summary | Summary By Day (incl. Month End Charge row) |
| Card_Type | items/amounts gross, refunds, net per card brand |
| Location | Summary By Location (per MID) |
| Paid_by_Others, Disputes | detail lines |
| Fee_Summary | Fees / Interchange / Service Charges by brand |
| Fee_Detail | every fee line with category and product carried down, plus basis amount, rate, item count parsed from the description |
| Fee_Category_Totals | printed category subtotals |
| Interchange | qualification lines (brand, sales, % sales, txns, rate, per-item, charge, recalculated charge) |
| Interchange_Brand_Totals, Amounts_Funded, Sales_Trend_13M | as printed |
| Unparsed | any row the parser could not map (should be empty) |

## How it works

* Text is read with PyMuPDF at character level and grouped into rows by baseline and columns by x position.
  Characters covered by a filled shape drawn on top (masked / redacted areas) are dropped and reported.
* A small state machine switches parser per section title; section context carries across page breaks.
* Money is handled as `Decimal`, so totals reconcile to the cent.
