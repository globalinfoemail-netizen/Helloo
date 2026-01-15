# GPSE Hub v2+ (Mock-ready)

## Run (Windows PowerShell)
```powershell
cd gpse_hub_v2_plus
python -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install -r requirements.txt
python app.py
```

Open:
- http://127.0.0.1:5000/dashboard
- http://127.0.0.1:5000/kpi-library
- http://127.0.0.1:5000/utilities
- http://127.0.0.1:5000/assistant

## Mock data
On first run a SQLite DB `gpse.db` is created and seeded with:
- departments
- KPI definitions (kpi_master) including SEV2/CRI/MTTR examples
- reports + versions + KPI snapshots

## Power BI (local demo)
Use Get Data → Web:
- http://127.0.0.1:5000/api/kpis
- http://127.0.0.1:5000/api/reports
- http://127.0.0.1:5000/api/kpi_master
