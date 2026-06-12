# HOA Quarterly Billing System

This folder contains a focused Excel billing workbook for quarterly HOA or condominium assessment billing.

## Main artifact

Download and review:

- `HOA_Quarterly_Billing_System.xlsx`

The workbook is intentionally macro-free for the first review so it opens safely in Excel, Google Sheets, Apple Numbers, and mobile spreadsheet viewers. It includes a `Macro_Blueprint` sheet describing the automation that should be added later if the layout is approved.

## Included sheets

- `Instructions` - workflow overview
- `Setup` - billing year, quarter, dates, dues, late fees, payment instructions, and invoice memo
- `Units_Owners` - owner/unit billing directory
- `Quarterly_Billing_Run` - invoice rows and billing formulas
- `Payments` - payment entry and receipt tracking
- `Ledger` - running debit/credit history
- `Aging_Report` - open balance and days-past-due review
- `Invoice_Template` - printable invoice view by invoice number
- `Owner_Statement` - statement placeholder for owner account history
- `Print_Email_Log` - delivery tracking
- `Macro_Blueprint` - recommended future VBA buttons and behavior
- `Lookups` - dropdown values

## Recommended use

1. Replace the SAMPLE rows in `Units_Owners`.
2. Update `Setup` for the target billing quarter.
3. Review generated rows in `Quarterly_Billing_Run`.
4. Select an invoice number in `Invoice_Template` and print or save to PDF.
5. Enter received payments in `Payments`.
6. Review balances in `Aging_Report` and ledger activity in `Ledger`.

## Regenerating the workbook

From the repository root:

```bash
python3 -m pip install -r HOA_Quarterly_Billing/requirements.txt
python3 HOA_Quarterly_Billing/generate_billing_workbook.py
```

The script recreates `HOA_Quarterly_Billing/HOA_Quarterly_Billing_System.xlsx`.
