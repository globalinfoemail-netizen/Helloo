# GPSE Report Hub v2 (Advanced Demo)

## What this is
A local, manager-demo friendly GPSE Reporting Hub:
- Aesthetic dashboard (Bootstrap)
- Type-to-search + filters
- Report library + report detail pages
- Versioning + KPI snapshots
- KPI summary cards (Avg SLA, Total P1, Avg MTTR, Total Risks)
- Power BI-ready APIs (JSON) + CSV export
- One-click PPT generation (demo deck) per report

## Run
1) Install Python 3.10+ and ensure `python --version` works.
2) In this folder:
   - `python -m pip install -r requirements.txt`
   - `python app.py`
3) Open: http://127.0.0.1:5000

## Power BI
Use **Get Data → Web**:
- http://127.0.0.1:5000/api/kpis
- http://127.0.0.1:5000/api/reports

Or use CSV:
- http://127.0.0.1:5000/export/kpis.csv

## Notes
- Storage is URL-based and works with SharePoint, network drive, or any internal repository later.
- This demo uses no real bank data.
