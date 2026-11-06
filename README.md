# Oil & Lube Dip Check — Olive Grove Depot

Designed by First Bus Franchise · Engineering

| File | What it is |
|------|------------|
| `Oil_Lube_Dip_Tracker.xlsx` | One-tab weekday dip sheet (12 parts × dates), Nov 2026 – Dec 2027 |
| `dashboard.html` | Open in a browser, load the tracker (and optionally the GL Flex export) |
| `tools/build_tracker.py` | Regenerates a blank tracker (`python tools/build_tracker.py`) |

## Workbook
One tab, **Dips**: part numbers and names down the side, weekday dates across the top (weekends not listed).
Type each morning's dip in litres; leave blank if not dipped. Issues and goods in are not typed in — they come from the GL Flex export.

## Dashboard
Load the saved tracker with **Load Dip Tracker** and the depot GL Flex export with **Load GL Flex** (xlsx/csv).
The dashboard flags sharp falls/rises, totals issues & goods in per dip period (Monday includes the weekend) and shows
any movement GL Flex doesn't explain. TransactionType codes (I = issue, R = goods in by default) can be changed on the overview page.
Alert levels per part are in the `ALERTS` table near the top of the dashboard script.
Needs internet the first time to load SheetJS and Chart.js from cdnjs.
