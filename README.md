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
Open `dashboard.html` in a browser and load any of the three files (all three for the full picture):

| Button | File | Gives |
|--------|------|-------|
| Dip Tracker | `Oil_Lube_Dip_Tracker.xlsx` | Physical dips |
| GL Flex | Depot GL Flex export | Issues, goods in, returns, stock-take adjustments |
| Stock on Hand | Stock on hand report | Current system qty and value, used to draw the system stock line |

How the figures work:
- **Signs:** positive TransactionQty = stock out, negative = stock in. `I` = issue; `RECV`/`RREC`/`RETN` = goods in, net of returns;
  `STTK` = stock-take adjustment, which changes the system figure only. You can change these on the overview page.
- **Dip periods:** a dip is compared with GL lines posted after the previous dip, up to and including the dip day. Weekend postings count towards Monday.
- **Sharp fall:** a fall bigger than 1.5× normal GL usage for the days covered (Monday covers 3 days), and never smaller than the part's minimum.
- **Variance:** dip change − (goods in − issued), shown in litres and in £ at the stock-on-hand average price.
- **System stock:** the stock-on-hand qty, worked back through GL movements. The "as at" date is read from the report's file name
  (e.g. `…2026-11-06T122407…`) and can be changed on the overview page.

Minimum alert levels and tolerances per part are in the `PARTS` table at the top of the dashboard script.
Needs internet to load SheetJS and Chart.js from cdnjs.
