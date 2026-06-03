# HOA Master Database Template

This folder contains an Excel-compatible, all-in-one HOA and condominium association database template.

## Files

| File | Purpose |
| --- | --- |
| `HOA_Master_Database_Template.xlsx` | Editable Excel workbook for HOA operations. |
| `generate_hoa_workbook.py` | Python generator used to rebuild or customize the workbook. |
| `requirements.txt` | Python dependency needed by the generator. |

## Workbook coverage

The workbook is designed around common U.S. HOA and condominium management needs, including:

- Association profile, jurisdiction, management, and records custodian references
- Unit and property registry
- Owner directory, resident directory, emergency contacts, and property manager details
- Proof of ownership registration and verification
- Access card inventory and delivery records, including a control check for the two-card-per-condo rule
- Billing accounts, assessment schedules, charges, payments, ledger records, and delinquency tracking
- Governance records, board and committee roles, meetings, elections, and voting records
- Document index for governing documents, minutes, financials, contracts, insurance, legal records, correspondence, and retention tracking
- Violations, architectural requests, maintenance/work orders, vendors, insurance, reserves, assets, vehicles, pets, communications, and owner record requests
- Lookup lists, dropdowns, starter formulas, dashboard summaries, and a data dictionary
- A `Research Coverage` worksheet mapping common U.S. HOA database needs to the included workbook modules

## Recommended setup order

1. Complete `Association Profile`.
2. Add each condo/unit in `Units`.
3. Add legal ownership in `Owners`.
4. Add people and companies in `Directory`.
5. Add verified documents in `Ownership Proof`.
6. Set up `Billing Accounts` and `Assessment Schedule`.
7. Enter `Access Cards`, including delivery date and acknowledgement/receipt link.
8. Use operational sheets as needed: `Charges`, `Payments`, `Violations`, `Architectural Requests`, `Maintenance`, `Documents`, and `Record Requests`.

## Important workbook conventions

- Use stable IDs consistently across sheets, such as `U-001`, `OWN-001`, `CON-001`, `CARD-001`, `ACCT-001`, `CHG-001`, and `PAY-001`.
- Each worksheet uses an Excel table. Add new records directly below the starter row.
- Dropdown values are maintained in `Lookups`.
- Formula-driven checks are visible in `Dashboard`, `Owners`, `Charges`, `Payments`, `Billing Ledger`, `Delinquency Tracking`, and `Unit Compliance`.
- `Unit Compliance` flags any unit with more than two currently assigned access cards.

## Legal and privacy note

HOA and condominium record requirements vary by state and by governing documents. This template includes common U.S. record categories such as financial records, owner rosters, meeting minutes, governing documents, contracts, reserve records, violations, architectural records, and member record requests, but it is not legal advice. Confirm retention periods, inspection rights, privacy rules, collection procedures, and disclosure limits with qualified counsel or the association's licensed manager.

Because the workbook can contain personal information, billing data, ownership documents, and legal notes, store it in a restricted location and limit editing access to authorized board members, managers, or records custodians.

## Regenerating the workbook

From the repository root:

```bash
python3 -m pip install -r HOA_Database/requirements.txt
python3 HOA_Database/generate_hoa_workbook.py
```

The script recreates `HOA_Database/HOA_Master_Database_Template.xlsx`.
