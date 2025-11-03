# Fire Incident Database — Streamlit App

A Streamlit app to view, add, and manage fire incident data stored in an Excel workbook.

## Features
- Loads your Excel workbook with these sheets detected automatically:
  - `Incidents`, `Incident_Detail`, `Incident_Times`, `Incident_Personnel`, `Incident_Apparatus`, `Incident_Actions`
  - Lookup sheets: `Personnel`, `Apparatus`, `List_*` (IncidentTypes, UnitTypes, Priorities, Dispositions, Actions, States)
- Add/edit incidents and their related detail records
- Filter/search incidents
- Export current data back to Excel (download or overwrite the source file)

## Quickstart (local)

```bash
# 1) Create & activate a virtual environment (optional)
python -m venv .venv
source .venv/bin/activate  # Windows: .venv\Scripts\activate

# 2) Install dependencies
pip install -r requirements.txt

# 3) Put your workbook in the repo root named fire_incident_db.xlsx
#    (or change the FILE_PATH in app.py)

# 4) Run the app
streamlit run app.py
```

## Deploy to GitHub + Streamlit Community Cloud
1. Create a new GitHub repo and upload this folder.
2. In Streamlit Cloud, create a new app from your repo and branch.
3. Set the app file to `app.py`.
4. Upload your Excel workbook (`fire_incident_db.xlsx`) to the repo or to Streamlit's file manager, or point `FILE_PATH` to your storage.

## Workbook assumptions
- Primary key for an incident is `IncidentID` (integer).
- Related tables reference `IncidentID`.
- Date/time fields: `Date`, `Alarm`, `Arrival`, `Clear` (strings or datetimes).
- You can customize columns and validation in the code where marked.
