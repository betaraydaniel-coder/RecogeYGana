"""Generate a Hacienda Los Cabos style billing statement workbook.

This artifact is designed from the provided statement screenshots. It keeps the
orange section bars, green Hacienda title, red totals, yellow total boxes,
currency labels, and multiple statement variations while keeping the data in
editable source sheets.
"""

from __future__ import annotations

from copy import copy
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation


OUTPUT = Path(__file__).with_name("Hacienda_Los_Cabos_Billing_Statements.xlsx")

ORANGE = "ED7D31"
GREEN = "008B45"
MAROON = "9E2F2F"
YELLOW = "FFFF00"
RED = "FF0000"
BLUE = "0070C0"
LIGHT_GREEN = "E2F0D9"
LIGHT_GRAY = "F2F2F2"
WHITE = "FFFFFF"
BLACK = "000000"

THIN_ORANGE = Side(style="thin", color=ORANGE)
THIN_GRAY = Side(style="thin", color="D9D9D9")


CONDO_ROWS = [
    ["A4", "Flor Rodriguez", "Spanish", "Individual ES - Quarter", "Pesos", 5470.45, "", "A-4"],
    ["B2", "Sample Owner B2", "Spanish", "Individual ES - Quarter", "Pesos", 5470.45, "", "B-2"],
    ["B3", "Fernando Garza", "Spanish", "Individual ES - Quarter", "Pesos", 5470.45, "", "B-3"],
    ["B6", "Cristopher Sanchez", "Spanish", "Individual ES - Quarter", "Pesos", 5470.45, "", "B-6"],
    ["B8", "Miriam Ramirez & Leonardo Robledo", "Spanish", "Individual ES - Quarter", "Pesos", 5632.21, "", "B-8"],
    ["C5", "Charles Chan & Teresa Andersen", "English", "Individual EN - Quarter", "Pesos/USD", 5632.21, "USD balance also shown", "C-5"],
    ["A5", "Ron and Jackie Davis", "English", "Annual EN - Account Statement", "Pesos", 10183.83, "Credit account example", "A-5"],
    ["TR", "TR Real Estate Services", "English", "Manager Group Summary", "Pesos", 0, "Group statement recipient", "TR"],
    ["A1", "TR Real Estate Services", "English", "Manager Unit", "Pesos", 5470.45, "Managed unit", "A1"],
    ["B5", "TR Real Estate Services", "English", "Manager Unit", "Pesos", 5470.45, "Managed unit", "B5"],
    ["C1", "TR Real Estate Services", "English", "Manager Unit", "Pesos", 5470.45, "Managed unit", "C1"],
    ["C2", "TR Real Estate Services", "English", "Manager Unit", "Pesos", 5470.45, "Managed unit", "C2"],
    ["C3", "TR Real Estate Services", "English", "Manager Unit", "Pesos", 5470.45, "Managed unit", "C3"],
    ["C4", "TR Real Estate Services", "English", "Manager Unit", "Pesos", 5470.45, "Managed unit", "C4"],
    ["D1", "TR Real Estate Services", "English", "Manager Unit", "Pesos", 5470.45, "Managed unit", "D1"],
    ["E2", "TR Real Estate Services", "English", "Manager Unit", "Pesos", 5470.45, "Managed unit", "E2"],
    ["E4", "TR Real Estate Services", "English", "Manager Unit", "Pesos", 5470.45, "Managed unit", "E4"],
    ["E6", "TR Real Estate Services", "English", "Manager Unit", "Pesos", 10183.83, "Managed unit higher fee", "E6"],
    ["F1", "TR Real Estate Services", "English", "Manager Unit", "Pesos", 5513.67, "Managed unit special amount", "F1"],
    ["H4", "TR Real Estate Services", "English", "Manager Unit", "Pesos", 5470.45, "Managed unit", "H4"],
    ["J5", "TR Real Estate Services", "English", "Manager Unit", "Pesos", 5470.45, "Managed unit", "J5"],
    ["J6", "TR Real Estate Services", "English", "Manager Unit", "Pesos", 5470.45, "Managed unit", "J6"],
    ["K5", "TR Real Estate Services", "English", "Manager Unit", "Pesos", 5470.45, "Managed unit", "K5"],
    ["LOCAL COMERCIAL", "TR Real Estate Services", "English", "Manager Unit", "Pesos", 1700.49, "Commercial unit", "LOCAL COMERCIAL"],
]


HOA_ROWS = [
    ["A4", "Q2 2026", "abr-26", 5470.45, 0, -5470.45, 0],
    ["A4", "Q2 2026", "may-26", 5470.45, 0, 0, 5470.45],
    ["A4", "Q2 2026", "jun-26", 5470.45, 0, 0, 5470.45],
    ["B2", "Q2 2026", "abr-26", 5470.45, 0, -5470.45, 0],
    ["B2", "Q2 2026", "may-26", 5470.45, 0, 0, 0],
    ["B2", "Q2 2026", "jun-26", 5470.45, 0, 0, 5470.45],
    ["B3", "Q2 2026", "abr-26", 5470.45, 0, -5470.45, 0],
    ["B3", "Q2 2026", "may-26", 5470.45, 0, -5470.45, 0],
    ["B3", "Q2 2026", "jun-26", 5470.45, 0, 0, 5470.45],
    ["B6", "Q1 2026", "ene-26", 5470.45, 165, -5472.45, 163],
    ["B6", "Q1 2026", "feb-26", 5470.45, 165, 0, 5635.45],
    ["B6", "Q1 2026", "mar-26", 5470.45, 165, 0, 5635.45],
    ["B6", "Q2 2026", "abr-26", 5470.45, 165, 0, 5635.45],
    ["B6", "Q2 2026", "may-26", 5470.45, 165, 0, 5635.45],
    ["B6", "Q2 2026", "jun-26", 5470.45, 0, 0, 5470.45],
    ["B8", "Q2 2026", "Apr-26", 5632.21, 0, -5633.00, -0.79],
    ["B8", "Q2 2026", "May-26", 5632.21, 0, -5631.42, 0],
    ["B8", "Q2 2026", "Jun-26", 5632.21, 0, 0, 5632.21],
    ["C5", "Q3 2026", "Jul-26", 5632.21, 0, 0, 5632.21],
    ["C5", "Q3 2026", "Aug-26", 5632.21, 0, 0, 5632.21],
    ["C5", "Q3 2026", "Sep-26", 5632.21, 0, 0, 5632.21],
    ["A5", "Q1 2026", "Jan-26", 10183.83, 0, -25643.40, -15459.57],
    ["A5", "Q1 2026", "Feb-26", 10183.83, 0, -15459.57, -5275.74],
    ["A5", "Q1 2026", "Mar-26", 10183.83, 0, -5275.74, 4908.09],
    ["A5", "Q2 2026", "Apr-26", 10183.83, 0, -96562.56, -81740.64],
    ["A5", "Q2 2026", "May-26", 10183.83, 0, 0, -71286.81],
    ["A5", "Q2 2026", "Jun-26", 10183.83, 0, 0, -6102.98],
    ["A5", "Q3 2026", "Jul-26", 10183.83, 0, 0, -50919.15],
    ["A5", "Q3 2026", "Aug-26", 10183.83, 0, 0, -40735.32],
    ["A5", "Q3 2026", "Sep-26", 10183.83, 0, 0, -30551.49],
    ["A5", "Q4 2026", "Oct-26", 10183.83, 0, 0, -20367.66],
    ["A5", "Q4 2026", "Nov-26", 10183.83, 0, 0, -10183.83],
    ["A5", "Q4 2026", "Dec-26", 10183.83, 0, 0, 0],
]


WATER_ROWS = [
    ["A4", "may-26", "1521", 4, 348, 283, 283, "Pesos"],
    ["B2", "may-26", "15", 9, 783, 0, 783, "Pesos"],
    ["B3", "may-26", "6 - 15", 9, 783, 0, 664.45, "Pesos"],
    ["B6", "feb-26", "711", 6, 522, 0, 522, "Pesos"],
    ["B6", "mar-26", "717", 6, 522, 522, 1044, "Pesos"],
    ["B6", "abr-26", "725", 8, 696, 1044, 1740, "Pesos"],
    ["B6", "may-26", "732", 7, 609, 1740, 2349, "Pesos"],
    ["B8", "May-26", "1810", 19, 1653, 0, 1653, "Pesos"],
    ["C5", "Jun-26", "566 - 567", 1, 87, 0, 87, "Pesos"],
    ["A5", "Oct-25", "308 - 308", 0, 0, -848, -848, "Pesos"],
    ["A5", "Nov-25", "308", 0, 0, 0, 0, "Pesos"],
    ["A5", "Dec-25", "308", 0, 0, 0, 0, "Pesos"],
    ["A5", "Jan-26", "308", 0, 0, 0, 0, "Pesos"],
    ["A5", "Feb-26", "312", 4, 348, -1000, -1000, "Pesos"],
    ["A5", "Mar-26", "315", 3, 261, -739, -739, "Pesos"],
    ["A5", "Apr-26", "0", 0, 0, -739, 0, "Pesos"],
]


GROUP_UNITS = [
    ["A1", 5470.45, 5470.45, 5470.45, 0, 5470.45],
    ["B5", 5470.45, 5470.45, 5470.45, 0, 5470.45],
    ["C1", 5470.45, 5470.45, 5470.45, 0, 5470.45],
    ["C2", 5470.45, 5470.45, 5470.45, 0, 5470.45],
    ["C3", 5470.45, 5470.45, 5470.45, 0, 5470.45],
    ["C4", 5470.45, 5470.45, 5470.45, 0, 5470.45],
    ["D1", 5470.45, 5470.45, 5470.45, 0, 5470.45],
    ["E2", 5470.45, 5470.45, 5470.45, 0, 5470.45],
    ["E4", 5470.45, 5470.45, 5470.45, 0, 5470.45],
    ["E6", 10183.83, 10183.83, 10183.83, 0, 10183.83],
    ["F1", 5513.67, 5513.67, 5513.67, 0, 5513.67],
    ["H4", 5470.45, 5470.45, 5470.45, 0, 5470.45],
    ["J5", 5470.45, 5470.45, 5470.45, 0, 5470.45],
    ["J6", 5470.45, 5470.45, 5470.45, 0, 5470.45],
    ["K5", 5470.45, 5470.45, 5470.45, 0, 5470.45],
    ["LOCAL\nCOMERCIAL", 1827.19, 1827.19, 1700.49, 0, 1700.49],
]


GROUP_WATER = [
    ["A1", 651, 5, 435],
    ["B5", 527, 2, 174],
    ["C1", 514, 1, 87],
    ["C2", 912, 1, 87],
    ["C3", 170, 4, 348],
    ["C4", 869, 8, 696],
    ["D1", 609, 4, 348],
    ["D5", 492, 1, 87],
    ["E2", 1500, 7, 609],
    ["E4", 1061, 7, 609],
    ["E6", 342, 1, 87],
    ["F1", 1734, 6, 522],
    ["H4", 1560, 4, 348],
    ["H6", 1603, 7, 609],
    ["J5", 99899, 5, 435],
    ["J6", 360, 4, 348],
    ["K5", 635, 4, 348],
]


def main() -> None:
    wb = Workbook()
    wb.remove(wb.active)
    wb.properties.title = "Hacienda Los Cabos Billing Statements"
    wb.properties.subject = "HOA maintenance and water billing statements"
    wb.properties.creator = "Cursor Cloud Agent"

    create_guide(wb)
    create_setup(wb)
    create_source_sheets(wb)
    create_statement_control(wb)
    create_individual_es(wb)
    create_individual_en_annual(wb)
    create_individual_en_quarter(wb)
    create_group_hoa(wb)
    create_group_water(wb)
    create_combined_due(wb)
    create_format_map(wb)
    create_macro_roadmap(wb)
    finalize(wb)
    wb.save(OUTPUT)
    print(f"Created {OUTPUT}")


def create_guide(wb: Workbook) -> None:
    ws = wb.create_sheet("Guide")
    ws.sheet_view.showGridLines = False
    add_logo(ws, start_col=6, start_row=2)
    ws["A1"] = "Hacienda Los Cabos Billing Statement Workbook"
    ws["A1"].font = Font(bold=True, size=18, color=GREEN)
    ws["A3"] = "Designed from the statement variations you provided."
    ws["A3"].font = Font(italic=True)
    rows = [
        ("What changed", "This workbook now focuses on reproducing the Hacienda statement formats, not a generic HOA database."),
        ("Source data", "Edit Condo_Directory, HOA_Fees, Water_Billing, and Payments_Credits."),
        ("Printable statements", "Use Individual_ES_Qtr, Individual_EN_Annual, Individual_EN_Qtr, Group_HOA_Quarter, Group_Water_Quarter, and Combined_Due."),
        ("Colors preserved", "Orange section headers, green logo/title text, red totals, blue credits, and yellow total boxes match the provided examples."),
        ("Next automation step", "After approval, convert to .xlsm and add buttons for generating all statements, exporting PDFs, and applying late fees."),
    ]
    for row_idx, (label, value) in enumerate(rows, start=6):
        ws.cell(row=row_idx, column=1, value=label).font = Font(bold=True, color=GREEN)
        ws.cell(row=row_idx, column=2, value=value)
        ws.cell(row=row_idx, column=2).alignment = Alignment(wrap_text=True, vertical="top")
    set_widths(ws, [24, 90, 10, 10, 10, 18, 18, 18])


def create_setup(wb: Workbook) -> None:
    ws = wb.create_sheet("Setup")
    ws.sheet_view.showGridLines = False
    add_logo(ws, 6, 1)
    ws["A1"] = "Statement Setup"
    ws["A1"].font = Font(bold=True, size=16, color=GREEN)
    setup_rows = [
        ("Association Name", "Hacienda Los Cabos"),
        ("Statement Date English", "June 12, 2026"),
        ("Statement Date Spanish", "viernes, 12 de junio de 2026"),
        ("Billing Year", 2026),
        ("Default Quarter", "Q2 2026"),
        ("Maintenance Currency", "Pesos"),
        ("Water Rate per m3", 87),
        ("Manager Recipient", "TR Real Estate Services"),
        ("Manager HOA Period", "Second quarter maintenance"),
        ("Manager Water Period", "April 2026 water consumption"),
        ("Combined Notice Label", "TOTAL DUE"),
    ]
    for row_idx, (label, value) in enumerate(setup_rows, start=5):
        ws.cell(row=row_idx, column=1, value=label).font = Font(bold=True, color=GREEN)
        ws.cell(row=row_idx, column=2, value=value)
        ws.cell(row=row_idx, column=2).fill = PatternFill("solid", fgColor="FFF2CC")
    ws["D5"] = "Supported statement variations"
    ws["D5"].font = Font(bold=True, color=GREEN)
    variations = [
        "Individual Spanish quarterly statement",
        "Individual English quarterly statement",
        "Individual English annual account statement",
        "Grouped manager HOA quarterly summary",
        "Grouped manager water consumption summary",
        "Combined maintenance + water total due notice",
        "Credit/negative balance display",
        "Pesos and USD labels",
    ]
    for row_idx, value in enumerate(variations, start=6):
        ws.cell(row=row_idx, column=4, value=value)
    set_widths(ws, [30, 36, 8, 48, 18, 18, 18, 18])


def create_source_sheets(wb: Workbook) -> None:
    condo = wb.create_sheet("Condo_Directory")
    create_plain_table(
        condo,
        "Condo / Owner Directory",
        ["Unit", "Owner / Recipient", "Language", "Statement Format", "Currency", "Default HOA Fee", "Notes", "Display Unit"],
        CONDO_ROWS,
    )
    add_validation(condo, "C5:C200", '"Spanish,English"')
    add_validation(condo, "E5:E200", '"Pesos,USD,Pesos/USD"')
    money_column(condo, 6)

    hoa = wb.create_sheet("HOA_Fees")
    create_plain_table(
        hoa,
        "HOA / Maintenance Fees",
        ["Unit", "Quarter", "Period", "Cuota / HOA Fee", "Interest", "Balance / Previous Balance", "Total"],
        HOA_ROWS,
    )
    for col in [4, 5, 6, 7]:
        money_column(hoa, col)

    water = wb.create_sheet("Water_Billing")
    create_plain_table(
        water,
        "Water Billing",
        ["Unit", "Period", "Water Meter Reading", "m3 / Cubic Meters", "Consumption", "Balance / Previous Balance", "Total", "Currency"],
        WATER_ROWS,
    )
    for col in [5, 6, 7]:
        money_column(water, col)

    payments = wb.create_sheet("Payments_Credits")
    create_plain_table(
        payments,
        "Payments, Credits, and Adjustments",
        ["Date", "Unit", "Type", "Description", "Currency", "Amount", "Applied To", "Notes"],
        [
            ["2026-06-12", "A5", "Credit", "Credit pesos shown on annual statement", "Pesos", -739, "Water", "Blue credit example"],
            ["2026-06-12", "C5", "USD Reference", "USD equivalent or separate account balance", "USD", 1152.48, "Maintenance", "Dual currency example"],
        ],
    )
    money_column(payments, 6)


def create_statement_control(wb: Workbook) -> None:
    ws = wb.create_sheet("Statement_Control")
    ws.sheet_view.showGridLines = False
    add_logo(ws, 6, 1)
    ws["A1"] = "Statement Control"
    ws["A1"].font = Font(bold=True, size=16, color=GREEN)
    rows = [
        ("Selected Unit", "B3"),
        ("Selected Quarter", "Q2 2026"),
        ("Selected Language", "Spanish"),
        ("Selected Date", "=Setup!B3"),
        ("Manager Recipient", "=Setup!B12"),
        ("Manager HOA Total", "=SUM(Group_HOA_Quarter!F9:F24)"),
        ("Manager Water Total", "=SUM(Group_Water_Quarter!E9:E25)"),
    ]
    for row_idx, (label, value) in enumerate(rows, start=5):
        ws.cell(row=row_idx, column=1, value=label).font = Font(bold=True, color=GREEN)
        ws.cell(row=row_idx, column=2, value=value)
        ws.cell(row=row_idx, column=2).fill = PatternFill("solid", fgColor="FFF2CC")
    ws["D5"] = "Review targets from screenshots"
    ws["D5"].font = Font(bold=True, color=GREEN)
    targets = [
        "A4/B2/B3/B8 style Spanish statements",
        "B6 style with Q1 + Q2 and water carry-forward",
        "A5 annual English account statement with credits",
        "TR Real Estate Services group HOA summary",
        "TR water consumption and combined total due",
        "C5 English quarter with Pesos + USD total boxes",
    ]
    for row_idx, value in enumerate(targets, start=6):
        ws.cell(row=row_idx, column=4, value=value)
    set_widths(ws, [26, 28, 8, 58, 18, 18, 18])


def create_individual_es(wb: Workbook) -> None:
    ws = wb.create_sheet("Individual_ES_Qtr")
    ws.sheet_view.showGridLines = False
    add_logo(ws, 6, 1)
    ws["A7"] = "Estimado"
    ws["A8"] = '=IFERROR(INDEX(Condo_Directory!$B:$B,MATCH(Statement_Control!$B$5,Condo_Directory!$A:$A,0)),"")'
    ws["A8"].font = Font(bold=True)
    ws["C8"] = "Unidad:"
    ws["D8"] = '=IFERROR(INDEX(Condo_Directory!$H:$H,MATCH(Statement_Control!$B$5,Condo_Directory!$A:$A,0)),Statement_Control!$B$5)'
    ws["D8"].font = Font(bold=True)
    ws["F8"] = "=Setup!B7"
    ws["A9"] = "A continuacion encontrara el detallado de las cuotas de mantenimiento y adeudo de agua a la fecha:"
    merge(ws, "A9:H9")
    section_header(ws, 10, ["Periodo", "Cuota", "", "Intereses", "Saldo", "", "Total", ""])
    months = ["abr-26", "may-26", "jun-26"]
    for idx, month in enumerate(months, start=11):
        ws.cell(idx, 1, month)
        ws.cell(idx, 2, "$")
        ws.cell(idx, 3, f'=SUMIFS(HOA_Fees!$D:$D,HOA_Fees!$A:$A,Statement_Control!$B$5,HOA_Fees!$C:$C,$A{idx})')
        ws.cell(idx, 4, f'=SUMIFS(HOA_Fees!$E:$E,HOA_Fees!$A:$A,Statement_Control!$B$5,HOA_Fees!$C:$C,$A{idx})')
        ws.cell(idx, 5, f'=SUMIFS(HOA_Fees!$F:$F,HOA_Fees!$A:$A,Statement_Control!$B$5,HOA_Fees!$C:$C,$A{idx})')
        ws.cell(idx, 6, "$")
        ws.cell(idx, 7, f'=SUMIFS(HOA_Fees!$G:$G,HOA_Fees!$A:$A,Statement_Control!$B$5,HOA_Fees!$C:$C,$A{idx})')
    section_header(ws, 15, ["Agua", "Lectura medidor", "", "m3", "Consumo", "Saldo", "Total", ""])
    ws["A16"] = "may-26"
    ws["B16"] = '=IFERROR(INDEX(Water_Billing!$C:$C,MATCH(1,(Water_Billing!$A:$A=Statement_Control!$B$5)*(Water_Billing!$B:$B=$A$16),0)),"")'
    ws["D16"] = '=SUMIFS(Water_Billing!$D:$D,Water_Billing!$A:$A,Statement_Control!$B$5,Water_Billing!$B:$B,$A$16)'
    ws["E16"] = '=SUMIFS(Water_Billing!$E:$E,Water_Billing!$A:$A,Statement_Control!$B$5,Water_Billing!$B:$B,$A$16)'
    ws["F16"] = '=SUMIFS(Water_Billing!$F:$F,Water_Billing!$A:$A,Statement_Control!$B$5,Water_Billing!$B:$B,$A$16)'
    ws["G16"] = '=SUMIFS(Water_Billing!$G:$G,Water_Billing!$A:$A,Statement_Control!$B$5,Water_Billing!$B:$B,$A$16)'
    ws["A18"] = '="Total a pagar "&Statement_Control!$B$5'
    ws["A18"].font = Font(bold=True)
    ws["C18"] = '=SUM(G11:G13)+G16'
    ws["C18"].fill = PatternFill("solid", fgColor=YELLOW)
    ws["C18"].font = Font(bold=True, color=RED)
    ws["D18"] = "Pesos"
    ws["D18"].font = Font(bold=True, color=RED)
    format_statement_grid(ws, 1, 18, 8)
    money_ranges(ws, ["C11:C13", "D11:G13", "E16:G16", "C18:C18"])
    set_widths(ws, [18, 12, 14, 14, 14, 12, 16, 12])
    ws.print_area = "A1:H20"


def create_individual_en_annual(wb: Workbook) -> None:
    ws = wb.create_sheet("Individual_EN_Annual")
    ws.sheet_view.showGridLines = False
    add_logo(ws, 6, 1)
    ws["A7"] = "Dear,"
    ws["A8"] = "Ron and Jackie Davis"
    ws["A8"].font = Font(bold=True)
    ws["C8"] = "Unit"
    ws["D8"] = "A-5"
    ws["D8"].font = Font(bold=True)
    ws["G8"] = "=Setup!B6"
    ws["A9"] = "Below you will find your current account statement"
    quarters = [
        ("Q1 2026", ["Jan-26", "Feb-26", "Mar-26"]),
        ("Q2 2026", ["Apr-26", "May-26", "Jun-26"]),
        ("Q3 2026", ["Jul-26", "Aug-26", "Sep-26"]),
        ("Q4 2026", ["Oct-26", "Nov-26", "Dec-26"]),
    ]
    row = 10
    for quarter, months in quarters:
        section_header(ws, row, [quarter, "HOA FEE", "Interests", "Balance", "Total", "", ""])
        for month in months:
            row += 1
            ws.cell(row, 1, month)
            ws.cell(row, 2, f'=SUMIFS(HOA_Fees!$D:$D,HOA_Fees!$A:$A,"A5",HOA_Fees!$C:$C,$A{row})')
            ws.cell(row, 3, f'=SUMIFS(HOA_Fees!$E:$E,HOA_Fees!$A:$A,"A5",HOA_Fees!$C:$C,$A{row})')
            ws.cell(row, 4, f'=SUMIFS(HOA_Fees!$F:$F,HOA_Fees!$A:$A,"A5",HOA_Fees!$C:$C,$A{row})')
            ws.cell(row, 5, f'=SUMIFS(HOA_Fees!$G:$G,HOA_Fees!$A:$A,"A5",HOA_Fees!$C:$C,$A{row})')
        row += 1
    section_header(ws, row, ["Water", "Water meter", "Cubic meters", "Balance", "Total", "", ""])
    for water_row in range(row + 1, row + 8):
        source_row = water_row - row
        ws.cell(water_row, 1, f'=INDEX(Water_Billing!$B$14:$B$20,{source_row})')
        ws.cell(water_row, 2, f'=INDEX(Water_Billing!$C$14:$C$20,{source_row})')
        ws.cell(water_row, 3, f'=INDEX(Water_Billing!$D$14:$D$20,{source_row})')
        ws.cell(water_row, 4, f'=INDEX(Water_Billing!$F$14:$F$20,{source_row})')
        ws.cell(water_row, 5, f'=INDEX(Water_Billing!$G$14:$G$20,{source_row})')
    total_row = row + 9
    ws.cell(total_row, 2, "HOA FEES\n2027")
    ws.cell(total_row, 3, "=0")
    ws.cell(total_row + 1, 2, "WATER")
    ws.cell(total_row + 1, 3, "=-739")
    ws.cell(total_row + 3, 2, "TOTAL")
    ws.cell(total_row + 3, 3, "=-739")
    for r in [total_row, total_row + 1, total_row + 3]:
        ws.cell(r, 3).fill = PatternFill("solid", fgColor=YELLOW)
        ws.cell(r, 3).font = Font(bold=True, color=BLUE if ws.cell(r, 3).value == "=-739" else RED)
    ws.cell(total_row + 4, 3, "CREDIT PESOS")
    ws.cell(total_row + 4, 3).font = Font(size=8, color=BLUE)
    format_statement_grid(ws, 1, total_row + 4, 7)
    money_ranges(ws, ["B11:E43", f"C{total_row}:C{total_row+3}"])
    set_widths(ws, [18, 16, 16, 16, 18, 12, 18])
    ws.print_area = f"A1:G{total_row + 5}"


def create_individual_en_quarter(wb: Workbook) -> None:
    ws = wb.create_sheet("Individual_EN_Qtr")
    ws.sheet_view.showGridLines = False
    add_logo(ws, 5, 1)
    ws["A7"] = "Dear,"
    ws["A8"] = "Charles Chan & Teresa Andersen"
    ws["A8"].font = Font(bold=True)
    ws["C8"] = "Unit C-5"
    ws["F8"] = "=Setup!B6"
    ws["A9"] = "Below you will find your current account statement"
    section_header(ws, 10, ["Q3", "Maintenance\nfee", "Interests", "Previous\nbalance", "Total", ""])
    for idx, month in enumerate(["Jul-26", "Aug-26", "Sep-26"], start=11):
        ws.cell(idx, 1, month)
        ws.cell(idx, 2, f'=SUMIFS(HOA_Fees!$D:$D,HOA_Fees!$A:$A,"C5",HOA_Fees!$C:$C,$A{idx})')
        ws.cell(idx, 3, f'=SUMIFS(HOA_Fees!$E:$E,HOA_Fees!$A:$A,"C5",HOA_Fees!$C:$C,$A{idx})')
        ws.cell(idx, 4, f'=SUMIFS(HOA_Fees!$F:$F,HOA_Fees!$A:$A,"C5",HOA_Fees!$C:$C,$A{idx})')
        ws.cell(idx, 5, f'=SUMIFS(HOA_Fees!$G:$G,HOA_Fees!$A:$A,"C5",HOA_Fees!$C:$C,$A{idx})')
    section_header(ws, 15, ["Water", "Water meter\nreading", "Water\nConsumption", "Previous\nbalance", "Total", ""])
    ws["A16"] = "Jun-26"
    ws["B16"] = '=SUMIFS(Water_Billing!$C:$C,Water_Billing!$A:$A,"C5",Water_Billing!$B:$B,$A$16)'
    ws["C16"] = '=SUMIFS(Water_Billing!$D:$D,Water_Billing!$A:$A,"C5",Water_Billing!$B:$B,$A$16)'
    ws["D16"] = '=SUMIFS(Water_Billing!$F:$F,Water_Billing!$A:$A,"C5",Water_Billing!$B:$B,$A$16)'
    ws["E16"] = '=SUMIFS(Water_Billing!$G:$G,Water_Billing!$A:$A,"C5",Water_Billing!$B:$B,$A$16)'
    ws["A18"] = "TOTAL C5"
    ws["B18"] = "=SUM(E11:E13)+E16"
    ws["B18"].fill = PatternFill("solid", fgColor=YELLOW)
    ws["C18"] = "Pesos"
    ws["B19"] = "1152.48"
    ws["B19"].fill = PatternFill("solid", fgColor=YELLOW)
    ws["C19"] = "USD"
    for cell in ["B18", "B19"]:
        ws[cell].font = Font(bold=True, color=RED)
    format_statement_grid(ws, 1, 20, 6)
    money_ranges(ws, ["B11:E16", "B18:B19"])
    set_widths(ws, [18, 18, 18, 18, 18, 12])
    ws.print_area = "A1:F21"


def create_group_hoa(wb: Workbook) -> None:
    ws = wb.create_sheet("Group_HOA_Quarter")
    ws.sheet_view.showGridLines = False
    add_logo(ws, 5, 1)
    ws["A7"] = "Dear,"
    ws["A8"] = "=Setup!B12"
    ws["A8"].font = Font(bold=True)
    ws["F8"] = "=Setup!B6"
    ws["A9"] = "Please find 2026 second quarter statement for the HOA fees next:"
    section_header(ws, 10, ["Condo", "April", "May", "June", "Previous\nbalance", "Total Second\nQuarter"])
    for row_idx, values in enumerate(GROUP_UNITS, start=11):
        for col_idx, value in enumerate(values, start=1):
            ws.cell(row_idx, col_idx, value=value)
        ws.cell(row_idx, 1).font = Font(bold=True)
    total_row = 11 + len(GROUP_UNITS)
    ws.cell(total_row, 5, "TOTAL")
    ws.cell(total_row, 6, f"=SUM(F11:F{total_row-1})")
    ws.cell(total_row, 5).font = Font(italic=True, size=14)
    ws.cell(total_row, 6).font = Font(bold=True, color=RED, size=14)
    money_ranges(ws, [f"B11:F{total_row}"])
    format_statement_grid(ws, 1, total_row, 6)
    set_widths(ws, [18, 16, 16, 16, 18, 22])
    ws.print_area = f"A1:F{total_row+1}"


def create_group_water(wb: Workbook) -> None:
    ws = wb.create_sheet("Group_Water_Quarter")
    ws.sheet_view.showGridLines = False
    add_logo(ws, 5, 1)
    ws["A7"] = "Dear,"
    ws["A8"] = "=Setup!B12"
    ws["A8"].font = Font(bold=True)
    ws["E8"] = "=Setup!B6"
    ws["A9"] = "Please find April 2026 water consumption next:"
    section_header(ws, 10, ["Condo", "Water\nmeter", "Cubic\nmeters", "Consumption", "Total"])
    for row_idx, values in enumerate(GROUP_WATER, start=11):
        ws.cell(row_idx, 1, values[0]).font = Font(bold=True)
        ws.cell(row_idx, 2, values[1])
        ws.cell(row_idx, 3, values[2])
        ws.cell(row_idx, 4, values[3])
        ws.cell(row_idx, 5, values[3])
    total_row = 11 + len(GROUP_WATER)
    ws.cell(total_row, 4, "TOTAL")
    ws.cell(total_row, 5, f"=SUM(E11:E{total_row-1})")
    ws.cell(total_row, 4).font = Font(size=14)
    ws.cell(total_row, 5).font = Font(bold=True, color=RED, size=14)
    ws.cell(total_row + 2, 3, "Second quarter\nmaintenance")
    ws.cell(total_row + 2, 4, "=Group_HOA_Quarter!F27")
    ws.cell(total_row + 3, 3, "Total water consumption May\n2026")
    ws.cell(total_row + 3, 4, f"=E{total_row}")
    ws.cell(total_row + 5, 3, "TOTAL DUE")
    ws.cell(total_row + 5, 4, f"=D{total_row+2}+D{total_row+3}")
    for row in [total_row + 2, total_row + 3, total_row + 5]:
        ws.cell(row, 4).fill = PatternFill("solid", fgColor=YELLOW)
        ws.cell(row, 4).font = Font(bold=True, color=RED, size=13 if row == total_row + 5 else 11)
    money_ranges(ws, [f"D11:E{total_row}", f"D{total_row+2}:D{total_row+5}"])
    format_statement_grid(ws, 1, total_row + 5, 5)
    set_widths(ws, [14, 16, 18, 18, 18])
    ws.print_area = f"A1:E{total_row+7}"


def create_combined_due(wb: Workbook) -> None:
    ws = wb.create_sheet("Combined_Due")
    ws.sheet_view.showGridLines = False
    add_logo(ws, 5, 1)
    ws["A7"] = "Dear,"
    ws["A8"] = "=Setup!B12"
    ws["A8"].font = Font(bold=True)
    ws["E8"] = "=Setup!B6"
    ws["A10"] = "Second quarter maintenance"
    ws["C10"] = "=Group_HOA_Quarter!F27"
    ws["A11"] = "Total water consumption May 2026"
    ws["C11"] = "=Group_Water_Quarter!E28"
    ws["A13"] = "TOTAL DUE"
    ws["C13"] = "=C10+C11"
    for row in [10, 11, 13]:
        ws.cell(row, 1).font = Font(bold=True)
        ws.cell(row, 3).fill = PatternFill("solid", fgColor=YELLOW)
        ws.cell(row, 3).font = Font(bold=True, color=RED, size=14 if row == 13 else 11)
    ws["D13"] = "Pesos"
    ws["D13"].font = Font(bold=True, color=RED)
    money_ranges(ws, ["C10:C13"])
    set_widths(ws, [34, 8, 22, 12, 18])
    format_statement_grid(ws, 1, 15, 5)
    ws.print_area = "A1:E16"


def create_format_map(wb: Workbook) -> None:
    ws = wb.create_sheet("Statement_Format_Map")
    ws.sheet_view.showGridLines = False
    ws["A1"] = "Statement Format Map"
    ws["A1"].font = Font(bold=True, size=16, color=GREEN)
    headers = ["Format", "Use When", "Sheet", "Key Variations Supported"]
    rows = [
        ["Individual Spanish quarterly", "Owner receives one quarter plus water detail in Spanish", "Individual_ES_Qtr", "Unidad, Periodo/Cuota/Intereses/Saldo/Total, Agua, Pesos total box"],
        ["Individual English annual", "Owner needs full-year account statement", "Individual_EN_Annual", "Q1-Q4 sections, water ledger, credit pesos"],
        ["Individual English quarterly", "Owner needs one quarter in English", "Individual_EN_Qtr", "Maintenance fee, previous balance, water, Pesos and USD total boxes"],
        ["Manager HOA group", "Manager receives HOA fees for many condos", "Group_HOA_Quarter", "April/May/June columns, previous balance, total second quarter"],
        ["Manager water group", "Manager receives water for many condos", "Group_Water_Quarter", "Water meter, cubic meters, consumption, total, combined due block"],
        ["Combined due", "Summary notice combining maintenance and water", "Combined_Due", "Yellow total cells and red total due"],
    ]
    create_plain_table(ws, "Statement Format Map", headers, rows, start_row=3)


def create_macro_roadmap(wb: Workbook) -> None:
    ws = wb.create_sheet("Macro_Roadmap")
    ws.sheet_view.showGridLines = False
    ws["A1"] = "Recommended Macro Roadmap"
    ws["A1"].font = Font(bold=True, size=16, color=GREEN)
    headers = ["Button", "Purpose", "Behavior"]
    rows = [
        ["Build selected statement", "Populate the correct printable sheet for one unit", "Read Condo_Directory statement format, set Statement_Control, refresh formulas, and select the target template."],
        ["Generate all owner PDFs", "Produce statements for every active unit", "Loop through Condo_Directory, choose matching template, export print area to PDF, and log the file."],
        ["Generate manager package", "Create the TR Real Estate Services HOA/water package", "Export Group_HOA_Quarter, Group_Water_Quarter, and Combined_Due as a bundled PDF set."],
        ["Apply water rate", "Calculate water consumption from cubic meters", "Multiply m3 by Setup water rate, preserve manual overrides where needed."],
        ["Apply late fee/interests", "Add monthly interest/late fee entries", "Apply only when approved by HOA policy and avoid duplicate fees."],
    ]
    create_plain_table(ws, "Macro Roadmap", headers, rows, start_row=3)


def add_logo(ws, start_col: int, start_row: int) -> None:
    ws.cell(start_row, start_col, "HACIENDA")
    ws.cell(start_row + 1, start_col, "LOS CABOS")
    ws.cell(start_row, start_col).font = Font(bold=True, size=28, color=GREEN, name="Georgia")
    ws.cell(start_row + 1, start_col).font = Font(bold=True, size=22, color=GREEN, name="Georgia")
    ws.cell(start_row, start_col + 2, ")")
    ws.cell(start_row, start_col + 2).font = Font(size=44, color=MAROON)


def create_plain_table(ws, title_text, headers, rows, start_row=4) -> None:
    ws.sheet_view.showGridLines = False
    ws["A1"] = title_text
    ws["A1"].font = Font(bold=True, size=16, color=GREEN)
    section_header(ws, start_row, headers)
    for row_idx, row_values in enumerate(rows, start=start_row + 1):
        for col_idx, value in enumerate(row_values, start=1):
            ws.cell(row_idx, col_idx, value)
            ws.cell(row_idx, col_idx).alignment = Alignment(wrap_text=True, vertical="top")
    ws.freeze_panes = f"A{start_row + 1}"
    ws.auto_filter.ref = f"A{start_row}:{get_column_letter(len(headers))}{start_row + max(len(rows), 1)}"
    set_widths(ws, [max(14, min(30, len(str(h)) + 5)) for h in headers])


def section_header(ws, row, labels) -> None:
    for col_idx, label in enumerate(labels, start=1):
        cell = ws.cell(row, col_idx, label)
        cell.fill = PatternFill("solid", fgColor=ORANGE)
        cell.font = Font(bold=True, color=WHITE)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = Border(top=THIN_ORANGE, bottom=THIN_ORANGE, left=THIN_ORANGE, right=THIN_ORANGE)


def format_statement_grid(ws, min_col, max_row, max_col) -> None:
    for row in range(1, max_row + 1):
        for col in range(min_col, max_col + 1):
            cell = ws.cell(row, col)
            cell.border = Border(top=THIN_GRAY, bottom=THIN_GRAY, left=THIN_GRAY, right=THIN_GRAY)
            alignment = copy(cell.alignment)
            alignment.vertical = "center"
            alignment.wrap_text = True
            cell.alignment = alignment
    for row in range(10, max_row + 1):
        for col in range(min_col, max_col + 1):
            ws.cell(row, col).border = Border(top=THIN_ORANGE, bottom=THIN_ORANGE, left=THIN_ORANGE, right=THIN_ORANGE)
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0


def set_widths(ws, widths) -> None:
    for col_idx, width in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(col_idx)].width = width


def money_column(ws, col) -> None:
    for row in range(5, 500):
        ws.cell(row, col).number_format = '$#,##0.00;[Blue]-$#,##0.00;"-"'


def money_ranges(ws, ranges) -> None:
    for cell_range in ranges:
        for row in ws[cell_range]:
            for cell in row:
                cell.number_format = '$#,##0.00;[Blue]-$#,##0.00;"-"'
                if isinstance(cell.value, (int, float)) and cell.value > 0:
                    cell.font = Font(color=RED)


def add_validation(ws, cell_range, formula) -> None:
    dv = DataValidation(type="list", formula1=formula, allow_blank=True)
    ws.add_data_validation(dv)
    dv.add(cell_range)


def merge(ws, cell_range) -> None:
    ws.merge_cells(cell_range)


def finalize(wb: Workbook) -> None:
    for ws in wb.worksheets:
        max_col = max(ws.max_column, 6)
        max_row = max(ws.max_row, 20)
        for row in range(1, max_row + 1):
            for col in range(1, max_col + 1):
                cell = ws.cell(row, col)
                if cell.border == Border():
                    cell.border = Border(top=THIN_GRAY, bottom=THIN_GRAY, left=THIN_GRAY, right=THIN_GRAY)
                if row > 4 and row % 2 == 0 and cell.value is None:
                    cell.fill = PatternFill("solid", fgColor=LIGHT_GRAY)
        ws.sheet_view.showGridLines = True
    wb.active = 0


if __name__ == "__main__":
    main()
