# Oil & Lube Dip Check — Olive Grove Depot

Designed by First Bus Franchise · Engineering

| File | What it is |
|------|------------|
| `Oil_Lube_Dip_Tracker.xlsx` | Weekday dip tracker (12 bulk/drum/IBC parts), Nov 2026 – Dec 2027 |
| `dashboard.html` | Open in a browser, load the tracker (and optionally the GL Flex export) |
| `tools/build_tracker.py` | Regenerates a blank tracker (`python tools/build_tracker.py`) |

## Workbook sheets
- **Dips** – type the morning dip (litres) per part. Weekdays only; leave blank if not dipped.
- **Summary / Charts** – latest position, alert counts, level charts.
- **Daily Change** – change since the last recorded dip. Red = sharp fall, green = sharp rise.
- **GL Flex** – paste the depot GL Flex export at A2 (standard A→AG column order).
- **Issues / Goods In** – GL Flex quantities for each dip period (previous dip → dip day, so Monday includes the weekend).
- **Variance** – Daily Change − (Goods In − Issues). Red = unexplained loss, amber = unexplained gain.
- **Parts** – fall/rise alert levels and variance tolerance per part (edit the yellow cells).
- **Config** – TransactionType codes for Issue (`I`) and Goods In (`R` by default; check against your export).

## Dashboard
Load the saved tracker with **Load Dip Tracker**. GL Flex can be loaded separately with **Load GL Flex** (xlsx/csv),
or picked up from the tracker's GL Flex sheet. You can change which TransactionType codes count as issues or goods in on the overview page.
Needs internet the first time to load SheetJS and Chart.js from cdnjs.
