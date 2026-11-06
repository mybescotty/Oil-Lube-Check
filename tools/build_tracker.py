"""Builds Oil_Lube_Dip_Tracker.xlsx - one-tab weekday dip sheet for Olive Grove Depot.

Parts down the side, weekday dates across the top. Issues / goods in come from the
GL Flex export, loaded straight into dashboard.html.

Re-run to regenerate a blank tracker:  python tools/build_tracker.py [out.xlsx]
"""
import sys
from datetime import date, timedelta
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.utils import get_column_letter as L

OUT = sys.argv[1] if len(sys.argv) > 1 else "Oil_Lube_Dip_Tracker.xlsx"
START, END = date(2026, 11, 2), date(2027, 12, 31)
PARTS = [
    ("100000", "Fuel Diesel Ultra Low Sulphur (L)"),
    ("207427", "Oil Transmission BULK (L)"),
    ("207429", "Oil Transmission 208L (L)"),
    ("207419", "Oil Axle 85W-140 BULK (L)"),
    ("207384", "Oil Engine 10W-30 BULK (L)"),
    ("274889", "Coolant Antifreeze Green Euro 6 IBC (L)"),
    ("213285", "Coolant Antifreeze Standard Red 50:50 1000L IBC (L)"),
    ("222636", "Oil Axle 85W-140 208L (L)"),
    ("241877", "Oil Compressor Hydrovane Fluidforce QCIM (L)"),
    ("256374", "Fluid Screenwash Premium 1000L (LTS)"),
    ("264210", "Oil Transmission 208L (L)"),
    ("264552", "Oil Engine 5W-30 208L (L)"),
]
NAVY = "2D2A6E"
F = "Arial"
thin = Side(style="thin", color="E0E0EC")
box = Border(left=thin, right=thin, top=thin, bottom=thin)
monday = Border(left=Side(style="medium", color=NAVY), right=thin, top=thin, bottom=thin)

dates, d = [], START
while d <= END:
    if d.weekday() < 5:
        dates.append(d)
    d += timedelta(days=1)

wb = Workbook()
ws = wb.active
ws.title = "Dips"
head = Font(name=F, size=10, bold=True, color="FFFFFF")
fill_head = PatternFill("solid", fgColor=NAVY)
fill_in = PatternFill("solid", fgColor="FFFDE7")

for c, h in ((1, "Part"), (2, "PartDesc")):
    x = ws.cell(1, c, h); x.font = head; x.fill = fill_head
    x.alignment = Alignment(vertical="center")
for j, dt in enumerate(dates):
    c = 3 + j
    x = ws.cell(1, c, dt); x.font = head; x.fill = fill_head; x.number_format = "ddd dd/mm/yy"
    x.alignment = Alignment(horizontal="center", vertical="center", text_rotation=90)
    ws.column_dimensions[L(c)].width = 8.5
for i, (pn, desc) in enumerate(PARTS):
    r = 2 + i
    ws.cell(r, 1, pn).font = Font(name=F, size=10, bold=True, color=NAVY)
    ws.cell(r, 2, desc).font = Font(name=F, size=10, color=NAVY)
    for c in (1, 2):
        ws.cell(r, c).border = box
    for j, dt in enumerate(dates):
        x = ws.cell(r, 3 + j)
        x.font = Font(name=F, size=10, color="0000FF"); x.fill = fill_in
        x.number_format = "#,##0"; x.border = monday if dt.weekday() == 0 else box

last_row, last_col = 1 + len(PARTS), L(2 + len(dates))
dv = DataValidation(type="decimal", operator="greaterThanOrEqual", formula1="0", allow_blank=True,
                    showErrorMessage=True, errorTitle="Dip reading", error="Enter litres as a number (0 or more).")
ws.add_data_validation(dv)
dv.add(f"C2:{last_col}{last_row}")

note = last_row + 2
ws.cell(note, 1, "Enter each weekday morning's dip in litres (numbers only). Leave blank if not dipped. "
                 "Weekends are not listed; bank holidays can be left blank.").font = Font(name=F, size=9, italic=True, color="6B6B8A")
ws.cell(note + 1, 1, "Issues & goods in come from the GL Flex export — load both files into dashboard.html. "
                     "Designed by First Bus Franchise · Engineering").font = Font(name=F, size=9, italic=True, color="6B6B8A")

ws.column_dimensions["A"].width = 10
ws.column_dimensions["B"].width = 46
ws.row_dimensions[1].height = 70
ws.freeze_panes = "C2"
ws.sheet_properties.tabColor = "E5007D"
wb.save(OUT)
print("saved", OUT, len(dates), "weekdays")
