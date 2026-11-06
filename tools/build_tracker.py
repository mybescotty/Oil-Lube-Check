"""Builds Oil_Lube_Dip_Tracker.xlsx - weekday dip tracker for Olive Grove Depot.

Re-run to regenerate a blank tracker:  python tools/build_tracker.py [out.xlsx]
"""
import sys
from datetime import date, timedelta
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.formatting.rule import FormulaRule
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.chart import LineChart, Reference
from openpyxl.chart.axis import DateAxis
from openpyxl.utils import get_column_letter as L
from openpyxl.comments import Comment

OUT = sys.argv[1] if len(sys.argv) > 1 else "Oil_Lube_Dip_Tracker.xlsx"
START, END = date(2026, 11, 2), date(2027, 12, 31)

# part, description, uom, container, fall alert, rise alert, variance tolerance
PARTS = [
    ("100000", "Fuel Diesel Ultra Low Sulphur (L)", "L", "Bulk tank", 15000, 5000, 1000),
    ("207427", "Oil Transmission BULK (L)", "L", "Bulk tank", 300, 300, 50),
    ("207429", "Oil Transmission 208L (L)", "L", "208L drum", 150, 150, 20),
    ("207419", "Oil Axle 85W-140 BULK (L)", "L", "Bulk tank", 200, 200, 30),
    ("207384", "Oil Engine 10W-30 BULK (L)", "L", "Bulk tank", 500, 500, 75),
    ("274889", "Coolant Antifreeze Green Euro 6 IBC (L)", "L", "IBC", 300, 300, 50),
    ("213285", "Coolant Antifreeze Standard Red 50:50 1000L IBC (L)", "L", "1000L IBC", 300, 300, 50),
    ("222636", "Oil Axle 85W-140 208L (L)", "L", "208L drum", 100, 100, 20),
    ("241877", "Oil Compressor Hydrovane Fluidforce QCIM (L)", "L", "Container", 10, 10, 2),
    ("256374", "Fluid Screenwash Premium 1000L (LTS)", "LTS", "1000L IBC", 300, 300, 50),
    ("264210", "Oil Transmission 208L (L)", "L", "208L drum", 150, 150, 20),
    ("264552", "Oil Engine 5W-30 208L (L)", "L", "208L drum", 150, 150, 20),
]
GL_HEADERS = ["TransactionType", "PostingDate", "TransactionNumber", "TransactionLine", "FiscalYear",
              "Period", "FromCode", "ToCode", "PODocType", "CompanyCode", "OPCODesc", "CostCentre",
              "CCDesc", "PLImpact", "PLDesc", "BalanceSheetCode", "BalanceSheetDesc", "UOM",
              "TransactionQty", "TransactionValue", "Plant", "Store", "Part", "PartDesc", "MtlGroup",
              "MtlGroupDesc", "FleetNumber", "PO", "POLine", "BIN", "Originator", "WorkOrder", "POLineType"]

NAVY, PINK = "2D2A6E", "E5007D"
FONT = "Arial"
f_base = Font(name=FONT, size=10)
f_head = Font(name=FONT, size=10, bold=True, color="FFFFFF")
f_input = Font(name=FONT, size=10, color="0000FF")
f_link = Font(name=FONT, size=10, color="008000")
fill_head = PatternFill("solid", fgColor=NAVY)
fill_sub = PatternFill("solid", fgColor="E0E0EC")
fill_input = PatternFill("solid", fgColor="FFFDE7")
fill_wkstart = PatternFill("solid", fgColor="F5F5F7")
thin = Side(style="thin", color="E0E0EC")
box = Border(left=thin, right=thin, top=thin, bottom=thin)
RED_FILL = PatternFill("solid", fgColor="F8D3CE")    # rgba(231,76,60,.12) on white approx
AMBER_FILL = PatternFill("solid", fgColor="FDE6C2")
GREEN_FILL = PatternFill("solid", fgColor="CDEBD9")
RED_FONT = Font(name=FONT, color="C0392B", bold=True)
GREEN_FONT = Font(name=FONT, color="27AE60", bold=True)

dates = []
d = START
while d <= END:
    if d.weekday() < 5:
        dates.append(d)
    d += timedelta(days=1)
FIRST = 3
LAST = FIRST + len(dates) - 1
NP = len(PARTS)
PC0 = 3  # first part column (C)
pcols = [L(PC0 + i) for i in range(NP)]

wb = Workbook()

# ---------------------------------------------------------------- Guide
g = wb.active
g.title = "Guide"
g.sheet_view.showGridLines = False
g.column_dimensions["A"].width = 3
g.column_dimensions["B"].width = 28
g.column_dimensions["C"].width = 100
g["B2"] = "Oil & Lube Dip Tracker — Olive Grove Depot"
g["B2"].font = Font(name=FONT, size=16, bold=True, color=NAVY)
g["B3"] = "Designed by First Bus Franchise · Engineering"
g["B3"].font = Font(name=FONT, size=9, color="6B6B8A")
rows = [
    ("HOW TO USE", None),
    ("1. Dips", "Enter each morning's dip reading (litres) against the date. Weekdays only — weekends are not listed. "
                "Leave a cell blank if a tank was not dipped; changes are measured from the last recorded dip."),
    ("2. GL Flex", "Paste the GL Flex export for the depot into the 'GL Flex' sheet starting at A2 (keep the header row). "
                   "Columns must stay in the standard order A→AG."),
    ("3. Parts", "Yellow cells are editable: Fall alert / Rise alert (litres between dips) and Variance tolerance."),
    ("4. Config", "Set the TransactionType codes that mean Issue and Goods In (defaults I and R)."),
    ("Daily Change", "Dip minus previous recorded dip. Red = sharp fall (≥ Fall alert). Green = sharp rise (≥ Rise alert)."),
    ("Issues / Goods In", "Totals from GL Flex for the period since the previous recorded dip up to and including the dip date "
                          "(so Monday picks up Saturday & Sunday transactions)."),
    ("Variance", "Daily Change − (Goods In − Issues). Negative = more product gone than was booked out (possible loss / "
                 "unbooked issue). Positive = more product than expected (unbooked delivery / dip error). "
                 "Red/amber when beyond the Variance tolerance."),
    ("Summary / Charts", "Latest position per part and a level chart for each part."),
    ("", None),
    ("CELL KEY", None),
    ("Blue text, yellow fill", "Input — type here"),
    ("Black text", "Formula — do not overwrite"),
    ("", None),
    ("EXAMPLE (format only)", None),
    ("Dips row", "Mon 09/11/2026 | 100000 → 126450 | 207427 → 2140 | 207429 → 416 | … (numbers only, no 'L')"),
    ("GL Flex row", "I | 26/05/2026 | 7601121 | 10 | FY 2027 | 03 | STOR | EVNT | | 1073 | First South Yorkshire Ltd | 2432 | "
                    "Olive Grove ENG | … | TransactionQty 20 | … | Part 207384"),
    ("", None),
    ("ASSUMPTIONS", None),
    ("Alert thresholds", "Starting values set by Engineering as a guide — tune them on the Parts sheet once a few weeks of dips are in."),
    ("Issue quantities", "Issues and Goods In use the absolute value of TransactionQty, so sign convention in the export doesn't matter."),
    ("Dates", f"Pre-filled weekdays {START:%d/%m/%Y} – {END:%d/%m/%Y}. Bank holidays are listed — just leave blank."),
]
r = 5
for k, v in rows:
    g.cell(r, 2, k)
    if v is None:
        g.cell(r, 2).font = Font(name=FONT, size=11, bold=True, color=PINK)
    else:
        g.cell(r, 2).font = Font(name=FONT, size=10, bold=True, color=NAVY)
        g.cell(r, 3, v).font = f_base
        g.cell(r, 3).alignment = Alignment(wrap_text=True, vertical="top")
        g.cell(r, 2).alignment = Alignment(vertical="top")
    r += 1
g["B16"].font = f_input; g["B16"].fill = fill_input

# ---------------------------------------------------------------- Parts
p = wb.create_sheet("Parts")
ph = ["Part", "PartDesc", "UOM", "Container", "Fall alert (L)", "Rise alert (L)", "Variance tolerance (L)", "Notes"]
for c, h in enumerate(ph, 1):
    x = p.cell(1, c, h); x.font = f_head; x.fill = fill_head
    x.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
for i, row in enumerate(PARTS):
    for c, v in enumerate(row, 1):
        x = p.cell(2 + i, c, v); x.font = f_base; x.border = box
        if c >= 5:
            x.font = f_input; x.fill = fill_input; x.number_format = "#,##0"
    p.cell(2 + i, 8).font = f_input; p.cell(2 + i, 8).fill = fill_input; p.cell(2 + i, 8).border = box
p.cell(1, 5).comment = Comment("Starting guide value set by Engineering — adjust once normal daily usage is known.", "Engineering")
for c, w in zip("ABCDEFGH", [10, 52, 7, 13, 13, 13, 14, 40]):
    p.column_dimensions[c].width = w
p.row_dimensions[1].height = 30
p.freeze_panes = "C2"

# ---------------------------------------------------------------- Config
cf = wb.create_sheet("Config")
cf.column_dimensions["A"].width = 30; cf.column_dimensions["B"].width = 12; cf.column_dimensions["C"].width = 70
for c, h in enumerate(["Setting", "Value", "Notes"], 1):
    x = cf.cell(1, c, h); x.font = f_head; x.fill = fill_head
cfg = [("Issue TransactionType", "I", "Code in GL Flex column A for stock issued out (e.g. to a work order)."),
       ("Goods In TransactionType", "R", "Code in GL Flex column A for goods received. Check your export and change if different.")]
for i, (a, b, c) in enumerate(cfg, 2):
    cf.cell(i, 1, a).font = f_base
    x = cf.cell(i, 2, b); x.font = f_input; x.fill = fill_input; x.border = box
    cf.cell(i, 3, c).font = f_base
ISSUE_T, GI_T = "Config!$B$2", "Config!$B$3"

# ---------------------------------------------------------------- GL Flex
gl = wb.create_sheet("GL Flex")
for c, h in enumerate(GL_HEADERS, 1):
    x = gl.cell(1, c, h); x.font = f_head; x.fill = fill_head
    gl.column_dimensions[L(c)].width = max(10, len(h) + 3)
gl.freeze_panes = "A2"
gl.auto_filter.ref = f"A1:{L(len(GL_HEADERS))}1"
GL = "'GL Flex'!"
GLR = 50000
g_type, g_date, g_qty, g_part = (f"{GL}${col}$2:${col}${GLR}" for col in ("A", "B", "S", "W"))


# ---------------------------------------------------------------- grid helper
def grid(name, title_fill=fill_head, width=13):
    ws = wb.create_sheet(name)
    hdr = [("Date", 14), ("Day", 6)]
    for c, (h, w) in enumerate(hdr, 1):
        x = ws.cell(1, c, h); x.font = f_head; x.fill = title_fill
        ws.cell(2, c).fill = fill_sub
        ws.column_dimensions[L(c)].width = w
    for i, pr in enumerate(PARTS):
        c = PC0 + i
        x = ws.cell(1, c, pr[0]); x.font = f_head; x.fill = title_fill
        x.alignment = Alignment(horizontal="center")
        y = ws.cell(2, c, f"=Parts!$B${2 + i}"); y.font = Font(name=FONT, size=8, color=NAVY)
        y.fill = fill_sub; y.alignment = Alignment(wrap_text=True, vertical="top", horizontal="center")
        ws.column_dimensions[L(c)].width = width
    ws.row_dimensions[2].height = 48
    for k, dt in enumerate(dates):
        rr = FIRST + k
        if name == "Dips":
            a = ws.cell(rr, 1, dt)
        else:
            a = ws.cell(rr, 1, f"=Dips!A{rr}")
        a.number_format = "dd/mm/yyyy"; a.font = f_base
        b = ws.cell(rr, 2, f'=TEXT(A{rr},"ddd")'); b.font = f_base
        if dt.weekday() == 0:
            a.fill = fill_wkstart; b.fill = fill_wkstart
    ws.freeze_panes = ws.cell(FIRST, PC0)
    return ws


def prev_rng(col, rr, sheet="Dips!"):
    return f"{sheet}{col}${FIRST}:{col}{rr - 1}"


def prev_date(col, rr):
    # date of last recorded dip before this row
    return f"LOOKUP(2,1/({prev_rng(col, rr)}<>\"\"),Dips!$A${FIRST}:$A{rr - 1})"


# ---------------------------------------------------------------- Dips (input)
dips = grid("Dips")
dv = DataValidation(type="decimal", operator="greaterThanOrEqual", formula1="0", allow_blank=True,
                    showErrorMessage=True, errorTitle="Dip reading", error="Enter litres as a number (0 or more).")
dips.add_data_validation(dv)
dv.add(f"{pcols[0]}{FIRST}:{pcols[-1]}{LAST}")
for rr in range(FIRST, LAST + 1):
    for col in pcols:
        x = dips[f"{col}{rr}"]; x.font = f_input; x.fill = fill_input; x.number_format = "#,##0"; x.border = box

# ---------------------------------------------------------------- Daily Change
chg = grid("Daily Change")
for i, col in enumerate(pcols):
    for rr in range(FIRST, LAST + 1):
        x = chg[f"{col}{rr}"]
        if rr == FIRST:
            x.value = None
        else:
            x.value = (f'=IF(Dips!{col}{rr}="","",IFERROR(Dips!{col}{rr}-'
                       f'LOOKUP(2,1/({prev_rng(col, rr)}<>""),{prev_rng(col, rr)}),""))')
        x.number_format = '+#,##0;-#,##0;0'; x.font = f_base
    rng = f"{col}{FIRST + 1}:{col}{LAST}"
    chg.conditional_formatting.add(rng, FormulaRule(
        formula=[f'AND(ISNUMBER({col}{FIRST + 1}),{col}{FIRST + 1}<=-Parts!$E${2 + i})'], fill=RED_FILL, font=RED_FONT))
    chg.conditional_formatting.add(rng, FormulaRule(
        formula=[f'AND(ISNUMBER({col}{FIRST + 1}),{col}{FIRST + 1}>=Parts!$F${2 + i})'], fill=GREEN_FILL, font=GREEN_FONT))


# ---------------------------------------------------------------- Issues / Goods In
def gl_sheet(name, ttype):
    ws = grid(name)
    for i, col in enumerate(pcols):
        for rr in range(FIRST + 1, LAST + 1):
            x = ws[f"{col}{rr}"]
            x.value = (f'=IF(Dips!{col}{rr}="","",IFERROR(ABS(SUMIFS({g_qty},{g_part},{col}$1,{g_type},{ttype},'
                       f'{g_date},">"&{prev_date(col, rr)},{g_date},"<"&($A{rr}+1))),""))')
            x.number_format = '#,##0;-#,##0;"-"'; x.font = f_base
    return ws


gl_sheet("Issues", ISSUE_T)
gl_sheet("Goods In", GI_T)

# ---------------------------------------------------------------- Variance
var = grid("Variance")
for i, col in enumerate(pcols):
    for rr in range(FIRST + 1, LAST + 1):
        x = var[f"{col}{rr}"]
        x.value = (f"=IF(OR('Daily Change'!{col}{rr}=\"\",Issues!{col}{rr}=\"\"),\"\","
                   f"'Daily Change'!{col}{rr}-('Goods In'!{col}{rr}-Issues!{col}{rr}))")
        x.number_format = '+#,##0;-#,##0;0'; x.font = f_base
    rng = f"{col}{FIRST + 1}:{col}{LAST}"
    c0 = f"{col}{FIRST + 1}"
    var.conditional_formatting.add(rng, FormulaRule(
        formula=[f'AND(ISNUMBER({c0}),{c0}<=-Parts!$G${2 + i})'], fill=RED_FILL, font=RED_FONT))
    var.conditional_formatting.add(rng, FormulaRule(
        formula=[f'AND(ISNUMBER({c0}),{c0}>=Parts!$G${2 + i})'], fill=AMBER_FILL))

# ---------------------------------------------------------------- Summary
s = wb.create_sheet("Summary", 1)
s.sheet_view.showGridLines = False
sh = ["Part", "PartDesc", "Last dip date", "Last dip (L)", "Last change (L)", "Largest fall (L)", "Largest rise (L)",
      "Fall alerts", "Rise alerts", "Total issues (L)", "Total goods in (L)", "Net variance (L)", "Variance alerts"]
for c, h in enumerate(sh, 1):
    x = s.cell(1, c, h); x.font = f_head; x.fill = fill_head
    x.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
s.row_dimensions[1].height = 32
for i, col in enumerate(pcols):
    rr = 2 + i
    D = f"Dips!{col}${FIRST}:{col}${LAST}"
    C = f"'Daily Change'!{col}${FIRST}:{col}${LAST}"
    V = f"Variance!{col}${FIRST}:{col}${LAST}"
    vals = [
        f"=Parts!A{rr}", f"=Parts!B{rr}",
        f'=IFERROR(LOOKUP(2,1/({D}<>""),Dips!$A${FIRST}:$A${LAST}),"")',
        f'=IFERROR(LOOKUP(2,1/({D}<>""),{D}),"")',
        f'=IFERROR(LOOKUP(2,1/({C}<>""),{C}),"")',
        f'=IF(COUNT({C})=0,"",MIN(0,MIN({C})))',
        f'=IF(COUNT({C})=0,"",MAX(0,MAX({C})))',
        f'=COUNTIF({C},"<="&-Parts!E{rr})',
        f'=COUNTIF({C},">="&Parts!F{rr})',
        f"=SUM(Issues!{col}${FIRST}:{col}${LAST})",
        f"=SUM('Goods In'!{col}${FIRST}:{col}${LAST})",
        f"=SUM({V})",
        f'=COUNTIF({V},"<="&-Parts!G{rr})+COUNTIF({V},">="&Parts!G{rr})',
    ]
    for c, v in enumerate(vals, 1):
        x = s.cell(rr, c, v); x.font = f_link if c <= 2 else f_base; x.border = box
        if c == 3:
            x.number_format = "ddd dd/mm/yyyy"
        elif c in (5, 6, 7, 12):
            x.number_format = '+#,##0;-#,##0;0'
        elif c >= 4:
            x.number_format = "#,##0"
    for c, rule_fill in ((8, RED_FILL), (9, GREEN_FILL), (13, AMBER_FILL)):
        cl = f"{L(c)}{rr}"
        s.conditional_formatting.add(cl, FormulaRule(formula=[f"{cl}>0"], fill=rule_fill))
for c, w in zip(range(1, 14), [9, 46, 15, 12, 12, 12, 12, 9, 9, 12, 13, 12, 10]):
    s.column_dimensions[L(c)].width = w
s.freeze_panes = "C2"

# ---------------------------------------------------------------- Charts
ch = wb.create_sheet("Charts", 2)
ch.sheet_view.showGridLines = False
ch["A1"] = "Dip level per part (updates as Dips are entered)"
ch["A1"].font = Font(name=FONT, size=12, bold=True, color=NAVY)
for i, col in enumerate(pcols):
    lc = LineChart()
    lc.title = f"{PARTS[i][0]} — {PARTS[i][1]}"
    lc.height, lc.width = 7, 16
    lc.y_axis.title = "Litres"
    lc.y_axis.numFmt = "#,##0"
    lc.legend = None
    data = Reference(dips, min_col=PC0 + i, min_row=FIRST, max_row=LAST)
    lc.add_data(data, titles_from_data=False)
    lc.set_categories(Reference(dips, min_col=1, min_row=FIRST, max_row=LAST))
    lc.x_axis = DateAxis(crossAx=100); lc.x_axis.number_format = "dd/mm"; lc.x_axis.majorTimeUnit = "days"
    lc.y_axis.crossAx = 500
    lc.x_axis.delete = False; lc.y_axis.delete = False
    srs = lc.series[0]
    srs.graphicalProperties.line.solidFill = NAVY
    srs.graphicalProperties.line.width = 20000
    srs.smooth = False
    lc.display_blanks = "span"
    ch.add_chart(lc, f"{'A' if i % 2 == 0 else 'K'}{3 + (i // 2) * 15}")

for ws in wb.worksheets:
    ws.sheet_properties.tabColor = PINK if ws.title in ("Dips", "GL Flex") else NAVY
wb.active = wb.sheetnames.index("Dips")
wb.calculation.fullCalcOnLoad = True
wb.save(OUT)
print("saved", OUT, len(dates), "weekdays")
