from openpyxl import Workbook
from openpyxl.styles import (PatternFill, Font, Alignment, Border, Side)
from openpyxl.styles.differential import DifferentialStyle
from openpyxl.formatting.rule import Rule, FormulaRule
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.utils import get_column_letter
from datetime import date

wb = Workbook()

# ── colour palette ────────────────────────────────────────────────────────────
BLUE       = "1a73e8"
WHITE      = "FFFFFF"
PAID_BG    = "b7e1cd"; PAID_FG    = "0f5132"
PEND_BG    = "fff3cd"; PEND_FG    = "856404"
OVER_BG    = "f8d7da"; OVER_FG    = "842029"
PART_BG    = "cce5ff"; PART_FG    = "004085"
CANC_BG    = "e2e3e5"; CANC_FG    = "383d41"
TOT_BG     = "d0e4f7"
ALT_ROW    = "f0f4ff"

def hdr_style():
    return Font(bold=True, color=WHITE, size=11)

def fill(hex_color):
    return PatternFill("solid", fgColor=hex_color)

def thin_border():
    s = Side(style="thin", color="CCCCCC")
    return Border(left=s, right=s, top=s, bottom=s)

# ═══════════════════════════════════════════════════════════════════════════════
# TAB 1 — MASTER DATA  (hidden reference sheet)
# ═══════════════════════════════════════════════════════════════════════════════
master = wb.active
master.title = "Master Data"

master.append(["Customer Name", "Company Name", "Phone Number", "", "Collectors", "", "Status"])
for cell in master[1]:
    cell.font      = Font(bold=True, color=WHITE)
    cell.fill      = fill(BLUE)
    cell.alignment = Alignment(horizontal="center")

sample_customers = [
    ["John Smith",  "ABC Corp",        "9876543210"],
    ["Jane Doe",    "XYZ Ltd",         "9123456789"],
    ["Raj Kumar",   "Raj Enterprises", "9000011111"],
]
for row in sample_customers:
    master.append(row + ["", "", "", ""])

collectors = ["Arun", "Sunita", "Vikram", "Priya"]
for i, c in enumerate(collectors, start=2):
    master.cell(row=i, column=5, value=c)

statuses = ["Pending", "Paid", "Overdue", "Partial", "Cancelled"]
for i, s in enumerate(statuses, start=2):
    master.cell(row=i, column=7, value=s)

master.column_dimensions["A"].width = 20
master.column_dimensions["B"].width = 20
master.column_dimensions["C"].width = 16
master.column_dimensions["E"].width = 14
master.column_dimensions["G"].width = 14
master.sheet_state = "visible"   # keep visible so user can add data

# ═══════════════════════════════════════════════════════════════════════════════
# TAB 2 — CRM DATA
# ═══════════════════════════════════════════════════════════════════════════════
crm = wb.create_sheet("CRM Data")

HEADERS = [
    "Customer Name", "Company Name", "Phone Number", "Status",
    "Bill Generate Date", "Amount", "Collected Amount", "Collection Date",
    "Collected By", "Pending Amount", "Delay Days",
    "Contacted Date", "Last Contacted Date", "Notes"
]
COL_WIDTHS = [20, 20, 16, 13, 18, 13, 18, 16, 16, 16, 12, 16, 20, 30]

crm.append(HEADERS)
for i, cell in enumerate(crm[1], start=1):
    cell.font      = Font(bold=True, color=WHITE, size=11)
    cell.fill      = fill(BLUE)
    cell.alignment = Alignment(horizontal="center", vertical="center")
    crm.column_dimensions[get_column_letter(i)].width = COL_WIDTHS[i-1]
crm.row_dimensions[1].height = 30
crm.freeze_panes = "A2"

ROWS = 100   # data rows

for r in range(2, ROWS + 2):
    # B: Company (VLOOKUP from Master Data)
    crm.cell(row=r, column=2).value  = f'=IFERROR(VLOOKUP(A{r},\'Master Data\'!$A:$C,2,0),"")'
    # C: Phone
    crm.cell(row=r, column=3).value  = f'=IFERROR(VLOOKUP(A{r},\'Master Data\'!$A:$C,3,0),"")'
    # J: Pending Amount
    crm.cell(row=r, column=10).value = f'=IF(F{r}="","",F{r}-G{r})'
    # K: Delay Days
    crm.cell(row=r, column=11).value = f'=IF(E{r}="","",IF(D{r}="Paid","",TODAY()-INT(E{r})))'

    # Date format: E, H, L, M
    for col in [5, 8, 12, 13]:
        crm.cell(row=r, column=col).number_format = "DD/MM/YYYY"
    # Currency: F, G, J
    for col in [6, 7, 10]:
        crm.cell(row=r, column=col).number_format = "#,##0.00"

    # Alternate row colour
    if r % 2 == 0:
        for col in range(1, len(HEADERS) + 1):
            cell = crm.cell(row=r, column=col)
            if not cell.fill or cell.fill.fgColor.rgb in ("00000000", "FFFFFFFF"):
                cell.fill = fill(ALT_ROW)

# ── Data Validations ──────────────────────────────────────────────────────────
# Status dropdown
dv_status = DataValidation(
    type="list",
    formula1='"Pending,Paid,Overdue,Partial,Cancelled"',
    allow_blank=True, showDropDown=False
)
dv_status.sqref = f"D2:D{ROWS+1}"
crm.add_data_validation(dv_status)

# Customer dropdown (from Master Data col A)
dv_customer = DataValidation(
    type="list",
    formula1="'Master Data'!$A$2:$A$51",
    allow_blank=True, showDropDown=False
)
dv_customer.sqref = f"A2:A{ROWS+1}"
crm.add_data_validation(dv_customer)

# Collector dropdown (from Master Data col E)
dv_collector = DataValidation(
    type="list",
    formula1="'Master Data'!$E$2:$E$20",
    allow_blank=True, showDropDown=False
)
dv_collector.sqref = f"I2:I{ROWS+1}"
crm.add_data_validation(dv_collector)

# ── Conditional Formatting ────────────────────────────────────────────────────
status_range = f"D2:D{ROWS+1}"
delay_range  = f"K2:K{ROWS+1}"

def status_rule(status, bg, fg):
    ds = DifferentialStyle(
        font=Font(color=fg, bold=False),
        fill=PatternFill(bgColor=bg)
    )
    return Rule(type="containsText", operator="containsText", text=status,
                dxf=ds, formula=[f'NOT(ISERROR(SEARCH("{status}",D2)))'])

crm.conditional_formatting.add(status_range, status_rule("Paid",      PAID_BG, PAID_FG))
crm.conditional_formatting.add(status_range, status_rule("Pending",   PEND_BG, PEND_FG))
crm.conditional_formatting.add(status_range, status_rule("Overdue",   OVER_BG, OVER_FG))
crm.conditional_formatting.add(status_range, status_rule("Partial",   PART_BG, PART_FG))
crm.conditional_formatting.add(status_range, status_rule("Cancelled", CANC_BG, CANC_FG))

# Delay > 30 → red
ds_red = DifferentialStyle(font=Font(color=OVER_FG), fill=PatternFill(bgColor=OVER_BG))
crm.conditional_formatting.add(delay_range,
    Rule(type="cellIs", operator="greaterThan", dxf=ds_red, formula=["30"]))

# Delay 1–30 → yellow
ds_yel = DifferentialStyle(font=Font(color=PEND_FG), fill=PatternFill(bgColor=PEND_BG))
crm.conditional_formatting.add(delay_range,
    Rule(type="cellIs", operator="between", dxf=ds_yel, formula=["1", "30"]))

# ═══════════════════════════════════════════════════════════════════════════════
# TAB 3 — CHARTS  (summary table; user adds charts manually in Google Sheets)
# ═══════════════════════════════════════════════════════════════════════════════
charts = wb.create_sheet("Charts")

charts["A1"].value      = "CRM Dashboard"
charts["A1"].font       = Font(bold=True, size=20, color=BLUE)

charts["A3"].value      = "By Status"
charts["A3"].font       = Font(bold=True, size=13)

chart_headers = ["Status", "Count", "Billed", "Collected", "Pending"]
for i, h in enumerate(chart_headers, start=1):
    cell = charts.cell(row=4, column=i, value=h)
    cell.font      = Font(bold=True, color=WHITE)
    cell.fill      = fill(BLUE)
    cell.alignment = Alignment(horizontal="center")

for i, st in enumerate(statuses, start=5):
    charts.cell(row=i, column=1, value=st)
    charts.cell(row=i, column=2).value = f'=COUNTIF(\'CRM Data\'!D:D,"{st}")'
    charts.cell(row=i, column=3).value = f'=SUMIF(\'CRM Data\'!D:D,"{st}",\'CRM Data\'!F:F)'
    charts.cell(row=i, column=4).value = f'=SUMIF(\'CRM Data\'!D:D,"{st}",\'CRM Data\'!G:G)'
    charts.cell(row=i, column=5).value = f'=SUMIF(\'CRM Data\'!D:D,"{st}",\'CRM Data\'!J:J)'
    for col in [3, 4, 5]:
        charts.cell(row=i, column=col).number_format = "#,##0.00"

# Totals
for i, col in enumerate(["A","B","C","D","E"], start=1):
    cell = charts.cell(row=10, column=i)
    cell.font = Font(bold=True)
    cell.fill = fill(TOT_BG)
    if i == 1:
        cell.value = "TOTAL"
    else:
        cell.value = f"=SUM({col}5:{col}9)"
        if i > 2:
            cell.number_format = "#,##0.00"

# Monthly trend
charts["A12"].value = "Monthly Trend (Last 6 Months)"
charts["A12"].font  = Font(bold=True, size=13)

month_headers = ["Month", "Billed", "Collected", "Pending"]
for i, h in enumerate(month_headers, start=1):
    cell = charts.cell(row=13, column=i, value=h)
    cell.font      = Font(bold=True, color=WHITE)
    cell.fill      = fill(BLUE)
    cell.alignment = Alignment(horizontal="center")

from datetime import date
import calendar

today = date.today()
for offset in range(5, -1, -1):
    mo  = (today.month - offset - 1) % 12 + 1
    yr  = today.year - ((today.month - offset - 1) // 12 + (1 if offset > today.month - 1 else 0))
    # Recalculate properly
    import datetime
    d = datetime.date(today.year, today.month, 1)
    # Go back 'offset' months
    month_num = today.month - offset
    year_num  = today.year
    while month_num <= 0:
        month_num += 12
        year_num  -= 1
    d = datetime.date(year_num, month_num, 1)
    row = 14 + (5 - offset)
    label = d.strftime("%b %Y")
    charts.cell(row=row, column=1, value=label)
    charts.cell(row=row, column=2).value = (
        f"=SUMPRODUCT((MONTH('CRM Data'!E$2:E$101)={month_num})"
        f"*(YEAR('CRM Data'!E$2:E$101)={year_num})*('CRM Data'!F$2:F$101))"
    )
    charts.cell(row=row, column=3).value = (
        f"=SUMPRODUCT((MONTH('CRM Data'!E$2:E$101)={month_num})"
        f"*(YEAR('CRM Data'!E$2:E$101)={year_num})*('CRM Data'!G$2:G$101))"
    )
    charts.cell(row=row, column=4).value = (
        f"=SUMPRODUCT((MONTH('CRM Data'!E$2:E$101)={month_num})"
        f"*(YEAR('CRM Data'!E$2:E$101)={year_num})*('CRM Data'!J$2:J$101))"
    )
    for col in [2, 3, 4]:
        charts.cell(row=row, column=col).number_format = "#,##0.00"

charts["A19"].value = "👆 Select A4:E9 and insert a Pie or Column chart in Google Sheets"
charts["A19"].font  = Font(italic=True, color="888888")
charts["A20"].value = "👆 Select A13:D19 and insert a Line chart for the monthly trend"
charts["A20"].font  = Font(italic=True, color="888888")

for col, w in zip(["A","B","C","D","E"], [20,16,16,16,16]):
    charts.column_dimensions[col].width = w

# ── Sheet order ───────────────────────────────────────────────────────────────
wb._sheets = [crm, charts, master]

# ── Save ──────────────────────────────────────────────────────────────────────
out = r"C:\Users\KGN\Desktop\Projects\Excel\CRM_Sheet.xlsx"
wb.save(out)
print(f"Saved: {out}")
