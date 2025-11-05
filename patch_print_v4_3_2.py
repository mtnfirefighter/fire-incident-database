# patch_print_v4_3_2.py
from typing import Dict
import pandas as pd
import streamlit.components.v1 as components

INCIDENT_EXTRAS = [
    "CallerName",
    "CallerPhone",
    "ReportWriter",  # manual
    "Approver",      # manual
]

def _ensure_columns(df: pd.DataFrame, cols):
    if df is None:
        df = pd.DataFrame()
    for c in cols:
        if c not in df.columns:
            df[c] = pd.NA
    return df

def apply_patch(env: dict):
    # Extend Incidents table with extras when ensured
    if "ensure_table" in env:
        orig = env["ensure_table"]
        def wrapped(data: Dict[str, pd.DataFrame], name: str, cols: list):
            if name == "Incidents":
                cols = list(cols) + [c for c in INCIDENT_EXTRAS if c not in cols]
            return orig(data, name, cols)
        env["ensure_table"] = wrapped

def _fetch_times_row(data: Dict[str, pd.DataFrame], pk: str, sel, ensure_columns):
    times_df = ensure_columns(data.get("Incident_Times", pd.DataFrame()), ["IncidentNumber","Alarm","Enroute","Arrival","Clear"])
    if not times_df.empty:
        match = times_df[times_df[pk].astype(str) == str(sel)]
        if not match.empty:
            return match.iloc[0].to_dict()
    return {}

def _get_incident_record(data: Dict[str, pd.DataFrame], pk: str, sel):
    rec_df = data.get("Incidents", pd.DataFrame())
    if rec_df.empty:
        return None
    rec = rec_df[rec_df[pk].astype(str) == str(sel)]
    if rec.empty:
        return None
    return rec.iloc[0].to_dict()

def render_incident_block(st, data: Dict[str, pd.DataFrame], PRIMARY_KEY: str, sel, ensure_columns):
    rec = _get_incident_record(data, PRIMARY_KEY, sel)
    if rec is None:
        st.warning("Incident not found.")
        return

    times_row = _fetch_times_row(data, PRIMARY_KEY, sel, ensure_columns)

    st.subheader(f"Incident {sel}")
    st.write(
        f"**Type:** {rec.get('IncidentType','')}  |  "
        f"**Priority:** {rec.get('ResponsePriority','')}  |  "
        f"**Alarm Level:** {rec.get('AlarmLevel','')}"
    )
    st.write(f"**Date:** {rec.get('IncidentDate','')}  **Time:** {rec.get('IncidentTime','')}")
    st.write(
        f"**Location:** {rec.get('LocationName','')} — "
        f"{rec.get('Address','')} {rec.get('City','')} {rec.get('State','')} {rec.get('PostalCode','')}"
    )

    # Caller / authorship
    caller_name  = rec.get('CallerName','')
    caller_phone = rec.get('CallerPhone','')
    writer_name  = rec.get('ReportWriter','')
    approver     = rec.get('Approver','')
    st.write(
        f"**Caller:** {caller_name if caller_name else '_N/A_'}  |  "
        f"**Caller Phone:** {caller_phone if caller_phone else '_N/A_'}"
    )
    st.write(
        f"**Report Written By:** {writer_name if writer_name else '_N/A_'}  |  "
        f"**Approved By:** {approver if approver else '_N/A_'}"
        f"{' — at ' + str(rec.get('ReviewedAt')) if rec.get('ReviewedAt') else ''}"
    )

    # Times
    st.write(
        f"**Times —** "
        f"Alarm: {times_row.get('Alarm','')}  |  "
        f"Enroute: {times_row.get('Enroute','')}  |  "
        f"Arrival: {times_row.get('Arrival','')}  |  "
        f"Clear: {times_row.get('Clear','')}"
    )

    # Narrative (read-only)
    st.write("**Narrative:**")
    st.text_area("Narrative (read-only)",
                 value=str(rec.get("Narrative","")),
                 height=220,
                 key=f"narrative_readonly_{sel}",
                 disabled=True)

    # Personnel & Apparatus
    ip = ensure_columns(data.get("Incident_Personnel", pd.DataFrame()), ["IncidentNumber","Name","Role","Hours","RespondedIn"])
    ia = ensure_columns(data.get("Incident_Apparatus", pd.DataFrame()), ["IncidentNumber","Unit","UnitType","Role","Actions"])
    ip_view = ip[ip[PRIMARY_KEY].astype(str) == str(sel)]
    ia_view = ia[ia[PRIMARY_KEY].astype(str) == str(sel)]

    st.markdown(f"**Personnel on Scene ({len(ip_view)}):**")
    person_cols = [c for c in ["Name","Role","Hours","RespondedIn"] if c in ip_view.columns]
    st.dataframe(ip_view[person_cols] if not ip_view.empty else ip_view, use_container_width=True, hide_index=True, key=f"grid_print_personnel_{sel}")

    st.markdown(f"**Apparatus on Scene ({len(ia_view)}):**")
    app_cols = [c for c in ["Unit","UnitType","Role","Actions"] if c in ia_view.columns]
    st.dataframe(ia_view[app_cols] if not ia_view.empty else ia_view, use_container_width=True, hide_index=True, key=f"grid_print_apparatus_{sel}")

def render_print_button(st, label: str = "🖨️ Print Report"):
    if st.button(label):
        components.html("<script>window.print()</script>", height=0, width=0)
