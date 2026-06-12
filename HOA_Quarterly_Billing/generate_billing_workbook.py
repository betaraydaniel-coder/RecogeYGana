"""Generate a focused HOA quarterly billing workbook.

The workbook is intentionally macro-free for review and compatibility. It uses
plain worksheets, formulas, dropdowns, and a printable invoice template. A
Macro_Blueprint sheet documents the recommended VBA automation to add after the
billing workflow is approved.
"""

from __future__ import annotations

from copy import copy
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.utils import get_column_letter


OUTPUT = Path(__file__).with_name("HOA_Quarterly_Billing_System.xlsx")

NAVY = "1F4E78"
BLUE = "D9EAF7"
GREEN = "70AD47"
LIGHT_GREEN = "E2F0D9"
GRAY = "808080"
LIGHT_GRAY = "F2F2F2"
WHITE = "FFFFFF"
YELLOW = "FFF2CC"


LOOKUPS = {
    "Quarter": ["Q1", "Q2", "Q3", "Q4"],
    "Billing Preference": ["Email", "Mail", "Both", "Do Not Bill"],
    "Yes/No": ["Yes", "No"],
    "Invoice Status": ["Draft", "Open", "Sent", "Paid", "Past Due", "Skip - Inactive", "Void"],
    "Payment Method": ["ACH", "Check", "Credit Card", "Debit Card", "Cash", "Money Order", "Wire", "Other"],
    "Receipt Sent": ["Yes", "No", "N/A"],
    "Ledger Type": ["Charge", "Payment", "Adjustment", "Late Fee", "Credit", "Write-Off"],
    "Delivery Status": ["Not Sent", "Sent", "Delivered", "Bounced", "Returned", "Printed", "Mailed"],
}


def main() -> None:
    wb = Workbook()
    wb.remove(wb.active)
    wb.properties.title = "HOA Quarterly Billing System"
    wb.properties.subject = "Quarterly HOA billing workbook"
    wb.properties.creator = "Cursor Cloud Agent"

    create_instructions(wb)
    create_setup(wb)
    create_lookups(wb)
    create_units_owners(wb)
    create_billing_run(wb)
    create_payments(wb)
    create_ledger(wb)
    create_aging_report(wb)
    create_invoice_template(wb)
    create_owner_statement(wb)
    create_print_email_log(wb)
    create_macro_blueprint(wb)
    finalize(wb)

    wb.save(OUTPUT)
    print(f"Created {OUTPUT}")


def create_instructions(wb: Workbook) -> None:
    ws = wb.create_sheet("Instructions")
    ws.sheet_view.showGridLines = False
    title(ws, "HOA Quarterly Billing System")
    ws["A2"] = "Focused Excel billing database for efficient quarterly HOA or condo assessment billing."
    ws["A2"].font = Font(italic=True, color=GRAY)

    sections = [
        (
            "Purpose",
            "This workbook is designed for one practical task: producing, tracking, and reviewing quarterly HOA billing more efficiently.",
        ),
        (
            "Recommended quarterly workflow",
            "1. Update Setup for the year, quarter, billing date, due date, and memo. 2. Update Units_Owners. 3. Review Quarterly_Billing_Run. 4. Use Invoice_Template to print/save invoices. 5. Enter payments in Payments. 6. Review Aging_Report and Ledger.",
        ),
        (
            "Macro strategy",
            "This file is macro-free for review and safe opening. The Macro_Blueprint sheet lists the automation I recommend after you approve the layout: generate billing, export PDFs, mark invoices sent, apply late fees, and create statements.",
        ),
        (
            "Important",
            "Rows marked SAMPLE are for review only. Replace or delete them before live billing. Review all formulas, amounts, late-fee rules, and collection practices against your HOA documents.",
        ),
    ]
    row = 4
    for heading, text in sections:
        ws.cell(row=row, column=1, value=heading).font = Font(bold=True, color=NAVY)
        ws.cell(row=row + 1, column=1, value=text).alignment = Alignment(wrap_text=True, vertical="top")
        ws.merge_cells(start_row=row + 1, start_column=1, end_row=row + 1, end_column=6)
        row += 3

    ws["A18"] = "Primary sheets"
    ws["A18"].font = Font(bold=True, color=NAVY)
    rows = [
        ("Setup", "Billing cycle control panel."),
        ("Units_Owners", "Billing directory and quarterly dues by unit."),
        ("Quarterly_Billing_Run", "Invoice rows generated from unit records."),
        ("Payments", "Payment entry and reference tracking."),
        ("Ledger", "Manual audit ledger for charges, payments, credits, and adjustments."),
        ("Invoice_Template", "Printable invoice view selected by invoice number."),
        ("Aging_Report", "Open balance and days-past-due review."),
        ("Macro_Blueprint", "Recommended automation for a later .xlsm version."),
    ]
    for r, (sheet, purpose) in enumerate(rows, start=19):
        ws.cell(row=r, column=1, value=sheet).font = Font(bold=True)
        ws.cell(row=r, column=2, value=purpose)
    ws.column_dimensions["A"].width = 28
    ws.column_dimensions["B"].width = 90


def create_setup(wb: Workbook) -> None:
    ws = wb.create_sheet("Setup")
    ws.sheet_view.showGridLines = False
    title(ws, "Billing Setup")
    labels = [
        ("Association Name", "Example HOA / Condominium Association"),
        ("Billing Year", 2026),
        ("Quarter", "Q1"),
        ("Billing Date", "1/1/2026"),
        ("Due Date", "1/31/2026"),
        ("Late Fee Date", "2/1/2026"),
        ("Standard Quarterly Dues", 350),
        ("Default Special Assessment", 0),
        ("Late Fee Amount", 25),
        ("Payment Instructions", "Pay by check, ACH, or approved online payment method."),
        ("Invoice Memo", "Quarterly assessment invoice. Please pay by the due date to avoid late fees."),
        ("Next Invoice Sequence", 1),
    ]
    for row, (label, value) in enumerate(labels, start=4):
        ws.cell(row=row, column=1, value=label).font = Font(bold=True, color=NAVY)
        ws.cell(row=row, column=2, value=value)
        ws.cell(row=row, column=2).fill = PatternFill("solid", fgColor=YELLOW)
        if "Date" in label:
            ws.cell(row=row, column=2).number_format = "m/d/yyyy"
        if "Amount" in label or "Dues" in label or "Assessment" in label:
            ws.cell(row=row, column=2).number_format = "$#,##0.00"
    add_list_validation(ws, "B6:B6", "Quarter")
    ws.column_dimensions["A"].width = 30
    ws.column_dimensions["B"].width = 70


def create_lookups(wb: Workbook) -> None:
    ws = wb.create_sheet("Lookups")
    ws.sheet_view.showGridLines = False
    title(ws, "Dropdown Values")
    for col, (name, values) in enumerate(LOOKUPS.items(), start=1):
        ws.cell(row=4, column=col, value=name)
        style_header(ws.cell(row=4, column=col))
        for row, value in enumerate(values, start=5):
            ws.cell(row=row, column=col, value=value)
        ws.column_dimensions[get_column_letter(col)].width = max(18, len(name) + 4)
    ws.freeze_panes = "A5"


def create_units_owners(wb: Workbook) -> None:
    ws = wb.create_sheet("Units_Owners")
    title(ws, "Units and Owners")
    headers = [
        "Unit ID",
        "Condo/Unit No.",
        "Owner Name",
        "Billing Email",
        "Phone",
        "Mailing Address",
        "Billing Preference",
        "Quarterly Dues",
        "Active?",
        "Prior Balance Override",
        "Notes",
    ]
    write_headers(ws, headers)
    samples = [
        ["U-001", "101", "SAMPLE Owner One", "owner1@example.com", "555-0101", "101 Main St", "Email", 350, "Yes", 0, "Replace sample row"],
        ["U-002", "102", "SAMPLE Owner Two", "owner2@example.com", "555-0102", "102 Main St", "Mail", 350, "Yes", 0, "Replace sample row"],
        ["U-003", "103", "SAMPLE Owner Three", "owner3@example.com", "555-0103", "103 Main St", "Both", 350, "Yes", 45, "Prior balance example"],
    ]
    for row, values in enumerate(samples, start=5):
        for col, value in enumerate(values, start=1):
            ws.cell(row=row, column=col, value=value)
    for row in range(8, 505):
        ws.cell(row=row, column=8, value="=Setup!$B$10")
        ws.cell(row=row, column=9, value="Yes")
        ws.cell(row=row, column=10, value=0)
    add_list_validation(ws, "G5:G504", "Billing Preference")
    add_list_validation(ws, "I5:I504", "Yes/No")
    set_widths(ws, [14, 16, 28, 30, 16, 38, 20, 18, 12, 24, 34])
    ws.freeze_panes = "A5"
    ws.auto_filter.ref = "A4:K504"


def create_billing_run(wb: Workbook) -> None:
    ws = wb.create_sheet("Quarterly_Billing_Run")
    title(ws, "Quarterly Billing Run")
    headers = [
        "Invoice No.",
        "Billing Year",
        "Quarter",
        "Unit ID",
        "Condo/Unit No.",
        "Owner Name",
        "Billing Email",
        "Charge Date",
        "Due Date",
        "Prior Balance",
        "Quarterly Dues",
        "Special Assessment",
        "Late Fee",
        "Total Due",
        "Payments Applied",
        "Balance Due",
        "Status",
        "Sent Date",
        "Notes",
    ]
    write_headers(ws, headers)
    for row in range(5, 505):
        idx = row - 4
        ws.cell(row=row, column=1, value=f'=IF($D{row}="","","INV-"&$B{row}&"-"&$C{row}&"-"&TEXT(ROW(A{idx}),"0000"))')
        ws.cell(row=row, column=2, value='=Setup!$B$5')
        ws.cell(row=row, column=3, value='=Setup!$B$6')
        ws.cell(row=row, column=4, value=f'=IFERROR(INDEX(Units_Owners!$A$5:$A$504,ROWS($A$5:A{row})),"")')
        ws.cell(row=row, column=5, value=f'=IF($D{row}="","",IFERROR(INDEX(Units_Owners!$B$5:$B$504,MATCH($D{row},Units_Owners!$A$5:$A$504,0)),""))')
        ws.cell(row=row, column=6, value=f'=IF($D{row}="","",IFERROR(INDEX(Units_Owners!$C$5:$C$504,MATCH($D{row},Units_Owners!$A$5:$A$504,0)),""))')
        ws.cell(row=row, column=7, value=f'=IF($D{row}="","",IFERROR(INDEX(Units_Owners!$D$5:$D$504,MATCH($D{row},Units_Owners!$A$5:$A$504,0)),""))')
        ws.cell(row=row, column=8, value='=Setup!$B$7')
        ws.cell(row=row, column=9, value='=Setup!$B$8')
        ws.cell(row=row, column=10, value=f'=IF($D{row}="","",IFERROR(INDEX(Units_Owners!$J$5:$J$504,MATCH($D{row},Units_Owners!$A$5:$A$504,0)),0)+SUMIFS(Ledger!$F:$F,Ledger!$B:$B,$D{row},Ledger!$D:$D,"<"&$H{row})-SUMIFS(Ledger!$G:$G,Ledger!$B:$B,$D{row},Ledger!$D:$D,"<"&$H{row}))')
        ws.cell(row=row, column=11, value=f'=IF($D{row}="","",IFERROR(INDEX(Units_Owners!$H$5:$H$504,MATCH($D{row},Units_Owners!$A$5:$A$504,0)),Setup!$B$10))')
        ws.cell(row=row, column=12, value='=IF($D5="","",Setup!$B$11)' if row == 5 else f'=IF($D{row}="","",Setup!$B$11)')
        ws.cell(row=row, column=13, value=f'=IF($D{row}="","",IF(AND(TODAY()>=Setup!$B$9,$P{row}>0),Setup!$B$12,0))')
        ws.cell(row=row, column=14, value=f'=IF($D{row}="","",SUM($J{row}:$M{row}))')
        ws.cell(row=row, column=15, value=f'=IF($A{row}="","",SUMIFS(Payments!$F:$F,Payments!$D:$D,$A{row}))')
        ws.cell(row=row, column=16, value=f'=IF($D{row}="","",MAX(0,$N{row}-$O{row}))')
        ws.cell(row=row, column=17, value=f'=IF($D{row}="","",IF(IFERROR(INDEX(Units_Owners!$I$5:$I$504,MATCH($D{row},Units_Owners!$A$5:$A$504,0)),"No")<>"Yes","Skip - Inactive",IF($P{row}=0,"Paid",IF(TODAY()>$I{row},"Past Due","Open"))))')
    add_list_validation(ws, "C5:C504", "Quarter")
    add_list_validation(ws, "Q5:Q504", "Invoice Status")
    set_widths(ws, [22, 14, 12, 14, 16, 28, 30, 16, 16, 18, 18, 20, 14, 18, 18, 18, 18, 16, 34])
    apply_money_formats(ws, [10, 11, 12, 13, 14, 15, 16])
    apply_date_formats(ws, [8, 9, 18])
    ws.freeze_panes = "A5"
    ws.auto_filter.ref = "A4:S504"


def create_payments(wb: Workbook) -> None:
    ws = wb.create_sheet("Payments")
    title(ws, "Payments")
    headers = [
        "Payment ID",
        "Date Received",
        "Unit ID",
        "Invoice No.",
        "Owner Name",
        "Amount Paid",
        "Payment Method",
        "Reference No.",
        "Deposit Account",
        "Receipt Sent?",
        "Notes",
    ]
    write_headers(ws, headers)
    for row in range(5, 505):
        ws.cell(row=row, column=5, value=f'=IF($C{row}="","",IFERROR(INDEX(Units_Owners!$C$5:$C$504,MATCH($C{row},Units_Owners!$A$5:$A$504,0)),""))')
    add_list_validation(ws, "G5:G504", "Payment Method")
    add_list_validation(ws, "J5:J504", "Receipt Sent")
    set_widths(ws, [16, 16, 14, 22, 28, 16, 18, 22, 22, 16, 40])
    apply_money_formats(ws, [6])
    apply_date_formats(ws, [2])
    ws.freeze_panes = "A5"
    ws.auto_filter.ref = "A4:K504"


def create_ledger(wb: Workbook) -> None:
    ws = wb.create_sheet("Ledger")
    title(ws, "Ledger")
    headers = ["Ledger ID", "Unit ID", "Invoice No.", "Transaction Date", "Type", "Debit", "Credit", "Balance", "Reference", "Notes"]
    write_headers(ws, headers)
    for row in range(5, 1005):
        ws.cell(row=row, column=8, value=f'=IF($B{row}="","",SUMIFS($F:$F,$B:$B,$B{row},$D:$D,"<="&$D{row})-SUMIFS($G:$G,$B:$B,$B{row},$D:$D,"<="&$D{row}))')
    add_list_validation(ws, "E5:E1004", "Ledger Type")
    set_widths(ws, [16, 14, 22, 18, 16, 14, 14, 14, 24, 44])
    apply_money_formats(ws, [6, 7, 8])
    apply_date_formats(ws, [4])
    ws.freeze_panes = "A5"
    ws.auto_filter.ref = "A4:J1004"


def create_aging_report(wb: Workbook) -> None:
    ws = wb.create_sheet("Aging_Report")
    title(ws, "Aging Report")
    headers = ["Unit ID", "Owner Name", "Invoice No.", "Due Date", "Balance Due", "Days Past Due", "Aging Bucket", "Status"]
    write_headers(ws, headers)
    for row in range(5, 505):
        ws.cell(row=row, column=1, value=f'=Quarterly_Billing_Run!D{row}')
        ws.cell(row=row, column=2, value=f'=Quarterly_Billing_Run!F{row}')
        ws.cell(row=row, column=3, value=f'=Quarterly_Billing_Run!A{row}')
        ws.cell(row=row, column=4, value=f'=Quarterly_Billing_Run!I{row}')
        ws.cell(row=row, column=5, value=f'=Quarterly_Billing_Run!P{row}')
        ws.cell(row=row, column=6, value=f'=IF($E{row}=0,0,MAX(0,TODAY()-$D{row}))')
        ws.cell(row=row, column=7, value=f'=IF($E{row}=0,"Current",IF($F{row}<=30,"1-30",IF($F{row}<=60,"31-60",IF($F{row}<=90,"61-90","90+"))))')
        ws.cell(row=row, column=8, value=f'=Quarterly_Billing_Run!Q{row}')
    set_widths(ws, [14, 28, 22, 16, 18, 18, 16, 18])
    apply_money_formats(ws, [5])
    apply_date_formats(ws, [4])
    ws.freeze_panes = "A5"
    ws.auto_filter.ref = "A4:H504"


def create_invoice_template(wb: Workbook) -> None:
    ws = wb.create_sheet("Invoice_Template")
    ws.sheet_view.showGridLines = False
    title(ws, "Invoice Template")
    ws["A3"] = "Invoice No."
    ws["B3"] = "INV-2026-Q1-0001"
    ws["B3"].fill = PatternFill("solid", fgColor=YELLOW)
    ws["A5"] = "=Setup!$B$4"
    ws["A5"].font = Font(bold=True, size=16, color=NAVY)
    rows = [
        ("Owner", '=IFERROR(INDEX(Quarterly_Billing_Run!$F:$F,MATCH($B$3,Quarterly_Billing_Run!$A:$A,0)),"")'),
        ("Unit", '=IFERROR(INDEX(Quarterly_Billing_Run!$E:$E,MATCH($B$3,Quarterly_Billing_Run!$A:$A,0)),"")'),
        ("Email", '=IFERROR(INDEX(Quarterly_Billing_Run!$G:$G,MATCH($B$3,Quarterly_Billing_Run!$A:$A,0)),"")'),
        ("Billing Year", '=IFERROR(INDEX(Quarterly_Billing_Run!$B:$B,MATCH($B$3,Quarterly_Billing_Run!$A:$A,0)),"")'),
        ("Quarter", '=IFERROR(INDEX(Quarterly_Billing_Run!$C:$C,MATCH($B$3,Quarterly_Billing_Run!$A:$A,0)),"")'),
        ("Charge Date", '=IFERROR(INDEX(Quarterly_Billing_Run!$H:$H,MATCH($B$3,Quarterly_Billing_Run!$A:$A,0)),"")'),
        ("Due Date", '=IFERROR(INDEX(Quarterly_Billing_Run!$I:$I,MATCH($B$3,Quarterly_Billing_Run!$A:$A,0)),"")'),
    ]
    for row, (label, formula) in enumerate(rows, start=7):
        ws.cell(row=row, column=1, value=label).font = Font(bold=True)
        ws.cell(row=row, column=2, value=formula)
    ws["A16"] = "Description"
    ws["B16"] = "Amount"
    style_header(ws["A16"])
    style_header(ws["B16"])
    items = [
        ("Prior Balance", '=IFERROR(INDEX(Quarterly_Billing_Run!$J:$J,MATCH($B$3,Quarterly_Billing_Run!$A:$A,0)),0)'),
        ("Quarterly Dues", '=IFERROR(INDEX(Quarterly_Billing_Run!$K:$K,MATCH($B$3,Quarterly_Billing_Run!$A:$A,0)),0)'),
        ("Special Assessment", '=IFERROR(INDEX(Quarterly_Billing_Run!$L:$L,MATCH($B$3,Quarterly_Billing_Run!$A:$A,0)),0)'),
        ("Late Fee", '=IFERROR(INDEX(Quarterly_Billing_Run!$M:$M,MATCH($B$3,Quarterly_Billing_Run!$A:$A,0)),0)'),
        ("Payments Applied", '=-IFERROR(INDEX(Quarterly_Billing_Run!$O:$O,MATCH($B$3,Quarterly_Billing_Run!$A:$A,0)),0)'),
        ("Balance Due", '=IFERROR(INDEX(Quarterly_Billing_Run!$P:$P,MATCH($B$3,Quarterly_Billing_Run!$A:$A,0)),0)'),
    ]
    for row, (label, formula) in enumerate(items, start=17):
        ws.cell(row=row, column=1, value=label)
        ws.cell(row=row, column=2, value=formula)
        ws.cell(row=row, column=2).number_format = "$#,##0.00"
    ws["A24"] = "Payment Instructions"
    ws["A24"].font = Font(bold=True)
    ws["A25"] = "=Setup!$B$13"
    ws["A27"] = "Invoice Memo"
    ws["A27"].font = Font(bold=True)
    ws["A28"] = "=Setup!$B$14"
    ws.merge_cells("A25:D25")
    ws.merge_cells("A28:D30")
    ws.column_dimensions["A"].width = 28
    ws.column_dimensions["B"].width = 24
    ws.column_dimensions["C"].width = 24
    ws.column_dimensions["D"].width = 24
    ws.print_area = "A1:D32"


def create_owner_statement(wb: Workbook) -> None:
    ws = wb.create_sheet("Owner_Statement")
    ws.sheet_view.showGridLines = False
    title(ws, "Owner Statement")
    ws["A3"] = "Unit ID"
    ws["B3"] = "U-001"
    ws["B3"].fill = PatternFill("solid", fgColor=YELLOW)
    ws["A5"] = "Owner"
    ws["B5"] = '=IFERROR(INDEX(Units_Owners!$C:$C,MATCH($B$3,Units_Owners!$A:$A,0)),"")'
    headers = ["Date", "Type", "Invoice No.", "Debit", "Credit", "Balance", "Notes"]
    write_headers(ws, headers, row=8)
    ws["A10"] = "Use Ledger filters by Unit ID for detailed statement history, or add a macro to populate this section automatically."
    ws.merge_cells("A10:G10")
    set_widths(ws, [16, 18, 22, 14, 14, 14, 44])


def create_print_email_log(wb: Workbook) -> None:
    ws = wb.create_sheet("Print_Email_Log")
    title(ws, "Print and Email Log")
    headers = ["Invoice No.", "Unit ID", "Owner Name", "Delivery Method", "Delivery Status", "Sent/Printed Date", "File Name or Email Reference", "Notes"]
    write_headers(ws, headers)
    add_list_validation(ws, "D5:D504", "Billing Preference")
    add_list_validation(ws, "E5:E504", "Delivery Status")
    set_widths(ws, [22, 14, 28, 20, 18, 20, 40, 40])
    apply_date_formats(ws, [6])
    ws.freeze_panes = "A5"
    ws.auto_filter.ref = "A4:H504"


def create_macro_blueprint(wb: Workbook) -> None:
    ws = wb.create_sheet("Macro_Blueprint")
    ws.sheet_view.showGridLines = False
    title(ws, "Macro Blueprint for Future .xlsm Version")
    headers = ["Button / Macro", "Purpose", "Recommended Behavior"]
    write_headers(ws, headers)
    rows = [
        (
            "Generate Quarterly Billing",
            "Create a clean billing run for the selected quarter.",
            "Clear prior draft rows after confirmation, loop through active Units_Owners rows, assign invoice numbers, copy dues/special assessment, calculate prior balance, and set status to Draft/Open.",
        ),
        (
            "Export Invoices to PDF",
            "Save one PDF per invoice from Invoice_Template.",
            "Loop through Quarterly_Billing_Run rows where status is Open or Sent, place invoice number into Invoice_Template!B3, export print area to PDF using file names like 2026_Q1_Unit_101_Owner.pdf.",
        ),
        (
            "Mark Billing Run as Sent",
            "Lock in the send/print step.",
            "Update Sent Date, status, and Print_Email_Log for all selected invoices; optionally prevent duplicate send without confirmation.",
        ),
        (
            "Apply Late Fees",
            "Add approved late fees after due date.",
            "For invoices still open after Setup!Late Fee Date, add the configured late fee once, update balance/status, and add a ledger row.",
        ),
        (
            "Generate Owner Statements",
            "Create statements by unit for owners with open balances.",
            "Filter Ledger by Unit ID, populate Owner_Statement, export or print statements for delinquency review.",
        ),
    ]
    for row, values in enumerate(rows, start=5):
        for col, value in enumerate(values, start=1):
            ws.cell(row=row, column=col, value=value)
            ws.cell(row=row, column=col).alignment = Alignment(wrap_text=True, vertical="top")
    ws["A12"] = "Note"
    ws["A12"].font = Font(bold=True, color=NAVY)
    ws["B12"] = "After you approve this workbook layout, the next step is to convert it to .xlsm and add these macros as real Excel buttons."
    ws.merge_cells("B12:C12")
    set_widths(ws, [30, 42, 90])


def title(ws, text: str) -> None:
    ws["A1"] = text
    ws["A1"].font = Font(bold=True, size=18, color=NAVY)


def write_headers(ws, headers, row: int = 4) -> None:
    for col, header in enumerate(headers, start=1):
        ws.cell(row=row, column=col, value=header)
        style_header(ws.cell(row=row, column=col))


def style_header(cell) -> None:
    cell.font = Font(bold=True, color=WHITE)
    cell.fill = PatternFill("solid", fgColor=NAVY)
    cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)


def add_list_validation(ws, cell_range: str, lookup_name: str) -> None:
    values = LOOKUPS[lookup_name]
    formula = '"' + ",".join(values) + '"'
    dv = DataValidation(type="list", formula1=formula, allow_blank=True)
    ws.add_data_validation(dv)
    dv.add(cell_range)


def set_widths(ws, widths) -> None:
    for idx, width in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(idx)].width = width


def apply_money_formats(ws, cols) -> None:
    for col in cols:
        for row in range(5, 1005):
            ws.cell(row=row, column=col).number_format = "$#,##0.00;[Red]-$#,##0.00;0"


def apply_date_formats(ws, cols) -> None:
    for col in cols:
        for row in range(5, 1005):
            ws.cell(row=row, column=col).number_format = "m/d/yyyy"


def finalize(wb: Workbook) -> None:
    thin_gray = Side(style="thin", color="D9E2F3")
    border = Border(left=thin_gray, right=thin_gray, top=thin_gray, bottom=thin_gray)
    for ws in wb.worksheets:
        max_row = min(max(ws.max_row, 30), 1005)
        max_col = max(ws.max_column, 4)
        for row in range(1, max_row + 1):
            for col in range(1, max_col + 1):
                cell = ws.cell(row=row, column=col)
                cell.border = border
                alignment = copy(cell.alignment)
                alignment.vertical = alignment.vertical or "top"
                cell.alignment = alignment
                if row > 4 and row % 2 == 0 and cell.value is None:
                    cell.fill = PatternFill("solid", fgColor=LIGHT_GRAY)
        ws.sheet_properties.pageSetUpPr.fitToPage = True
        ws.page_setup.fitToWidth = 1
        ws.page_setup.fitToHeight = 0
    wb.active = 0


if __name__ == "__main__":
    main()
