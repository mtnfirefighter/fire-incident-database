# patch_print_v4_3_2.py
from typing import Dict
import pandas as pd

INCIDENT_EXTRAS = [
    "CallerName",
    "CallerPhone",
    "ReportWriter",
    "Approver",
]

def _ensure_columns(df: pd.DataFrame, cols):
    if df is None:
        df = pd.DataFrame()
    for c in cols:
        if c not in df.columns:
            df[c] = pd.NA
    return df

def apply_patch(env: dict):
    if "ensure_table" in env:
        orig = env["ensure_table"]
        def wrapped(data: Dict[str, pd.DataFrame], name: str, cols: list):
            if name == "Incidents":
                cols = list(cols) + [c for c in INCIDENT_EXTRAS if c not in cols]
            return orig(data, name, cols)
        env["ensure_table"] = wrapped

def render_print_block(st, data: Dict[str, pd.DataFrame], PRIMARY_KEY: str, sel, ensure_columns):
    rec_df = data.get("Incidents", pd.DataFrame())
    if rec_df.empty:
        st.warning("No incidents found."); return
    rec = rec_df[rec_df[PRIMARY_KEY].astype(str) == str(sel)]
    if rec.empty:
        st.warning("Incident not found."); return
    rec = rec.iloc[0].to_dict()

    times_df = ensure_columns(data.get("Incident_Times", pd.DataFrame()), ["IncidentNumber","Alarm","Enroute","Arrival","Clear"])
    times_row = {}
    if not times_df.empty:
        match = times_df[times_df[PRIMARY KEY].astype(str) == str(sel)]
        if not match.empty:
            times_row = match.iloc[0].to_dict()

    st.subheader(f"Incident {sel}")
    st.write(f"**Type:** {rec.get('IncidentType','')}  |  **Priority:** {rec.get('ResponsePriority','')}  |  **Alarm Level:** {rec.get('AlarmLevel','')}")
    st.write(f"**Date:** {rec.get('IncidentDate','')}  **Time:** {rec.get('IncidentTime','')}")
    st.write(f"**Location:** {rec.get('LocationName','')} — {rec.get('Address','')} {rec.get('City','')} {rec.get('State','')} {rec.get('PostalCode','')}")

    caller_name = rec.get("CallerName",""); caller_phone = rec.get("CallerPhone","")
    report_writer = rec.get("ReportWriter",""); approver_name = rec.get("Approver","")
    st.write(f"**Caller:** {caller_name if caller_name else '_N/A_'}  |  **Caller Phone:** {caller_phone if caller_phone else '_N/A_'}")
    st.write(f"**Report Written By:** {report_writer if report_writer else '_N/A_'}  |  **Approved By:** {approver_name if approver_name else '_N/A_'}  {'at ' + str(rec.get('ReviewedAt')) if rec.get('ReviewedAt') else ''}")

    st.write(f"**Times —** Alarm: {times_row.get('Alarm','')}  |  Enroute: {times_row.get('Enroute','')}  |  Arrival: {times_row.get('Arrival','')}  |  Clear: {times_row.get('Clear','')}")

    st.write("**Narrative:**")
    st.text_area("Narrative (read-only)", value=str(rec.get("Narrative","")), height=220, key="narrative_readonly_print", disabled=True)

    ip = ensure_columns(data.get("Incident_Personnel", pd.DataFrame()), ["IncidentNumber","Name","Role","Hours","RespondedIn"])
    ia = ensure_columns(data.get("Incident_Apparatus", pd.DataFrame()), ["IncidentNumber","Unit","UnitType","Role","Actions"])
    ip_view = ip[ip[PRIMARY_KEY].astype(str) == str(sel)]
    ia_view = ia[ia[PRIMARY_KEY].astype(str) == str(sel)]

    st.markdown(f"**Personnel on Scene ({len(ip_view)}):**")
    person_cols = [c for c in ["Name","Role","Hours","RespondedIn"] if c in ip_view.columns]
    st.dataframe(ip_view[person_cols] if not ip_view.empty else ip_view, use_container_width=True, hide_index=True, key="grid_print_personnel")

    st.markdown(f"**Apparatus on Scene ({len(ia_view)}):**")
    app_cols = [c for c in ["Unit","UnitType","Role","Actions"] if c in ia_view.columns]
    st.dataframe(ia_view[app_cols] if not ia_view.empty else ia_view, use_container_width=True, hide_index=True, key="grid_print_apparatus")
