# HOA Quarterly Billing System

This folder contains a focused Excel billing workbook for quarterly HOA or condominium assessment billing.

## Main artifact

Download and review:

- `Hacienda_Los_Cabos_Billing_Statements.xlsx`
- `HOA_Quarterly_Billing_System.xlsx`

`Hacienda_Los_Cabos_Billing_Statements.xlsx` is the updated statement-focused workbook based on the provided Hacienda Los Cabos examples. It keeps the orange section bars, green title/logo treatment, red totals, yellow total boxes, credit styling, Pesos/USD labels, individual statements, manager/group summaries, water-consumption formats, and combined total-due notices.

The workbook is intentionally macro-free for the first review so it opens safely in Excel, Google Sheets, Apple Numbers, and mobile spreadsheet viewers. It includes a macro roadmap describing the automation that should be added later if the layout is approved.

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

The Hacienda-style workbook adds:

- `Condo_Directory`
- `HOA_Fees`
- `Water_Billing`
- `Payments_Credits`
- `Statement_Control`
- `Individual_ES_Qtr`
- `Individual_EN_Annual`
- `Individual_EN_Qtr`
- `Group_HOA_Quarter`
- `Group_Water_Quarter`
- `Combined_Due`
- `Statement_Format_Map`
- `Macro_Roadmap`

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
python3 HOA_Quarterly_Billing/generate_hacienda_statement_workbook.py
```

The scripts recreate `HOA_Quarterly_Billing/HOA_Quarterly_Billing_System.xlsx` and `HOA_Quarterly_Billing/Hacienda_Los_Cabos_Billing_Statements.xlsx`.
