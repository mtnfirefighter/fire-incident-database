
import os
import io
from datetime import datetime, date
from typing import Dict, List

import pandas as pd
import streamlit as st

st.set_page_config(page_title="Fire Incident DB", page_icon="🚒", layout="wide")

# ---------- Settings ----------
DEFAULT_FILE = os.path.join(os.path.dirname(__file__), "fire_incident_db.xlsx")
PRIMARY_KEY = "IncidentNumber"
CHILD_TABLES = {
    "Incident_Times": ["IncidentNumber","Alarm","Enroute","Arrival","Clear"],
    "Incident_Personnel": ["IncidentNumber","Name","Role","Hours"],
    "Incident_Apparatus": ["IncidentNumber","Unit","UnitType","Role","Actions"],
    "Incident_Actions": ["IncidentNumber","Action","Notes"],
}
PERSONNEL_SCHEMA = ["PersonnelID","Name","UnitNumber","Rank","Badge","Phone","Email","Address","City","State","PostalCode","Certifications","Active"]
APPARATUS_SCHEMA = ["ApparatusID","UnitNumber","CallSign","UnitType","GPM","TankSize","SeatingCapacity","Station","Active"]

LOOKUP_SHEETS = {
    "List_IncidentType": "IncidentType",
    "List_AlarmLevel": "AlarmLevel",
    "List_ResponsePriority": "ResponsePriority",
    "List_PersonnelRoles": "Role",
    "List_UnitTypes": "UnitType",
    "List_Actions": "Action",
    "List_States": "State",
}
DATE_LIKE = {"IncidentDate"}
TIME_LIKE = {"IncidentTime","Alarm","Enroute","Arrival","Clear"}

# ---------- Helpers ----------
def load_workbook(path: str) -> Dict[str, pd.DataFrame]:
    try:
        xls = pd.ExcelFile(path)
        data = {name: xls.parse(name) for name in xls.sheet_names}
        return data
    except Exception as e:
        st.error(f"Failed to load: {e}")
        return {}

def save_workbook_to_bytes(dfs: Dict[str, pd.DataFrame]) -> bytes:
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="xlsxwriter") as writer:
        for sheet, df in dfs.items():
            df.to_excel(writer, sheet_name=sheet, index=False)
    buf.seek(0)
    return buf.read()

def get_lookups(data: Dict[str, pd.DataFrame]) -> Dict[str, List[str]]:
    out: Dict[str, List[str]] = {}
    for sheet, col in LOOKUP_SHEETS.items():
        if sheet in data and not data[sheet].empty:
            header = data[sheet].columns[0]
            out[col] = data[sheet][header].dropna().astype(str).tolist()
    return out

def ensure_columns(df: pd.DataFrame, cols: List[str]) -> pd.DataFrame:
    for c in cols:
        if c not in df.columns:
            df[c] = pd.NA
    return df

def render_dynamic_form(df: pd.DataFrame, lookups: Dict[str, List[str]], defaults: dict) -> dict:
    vals = {}
    prefer_first = [PRIMARY_KEY, "IncidentDate", "IncidentTime", "IncidentType", "ResponsePriority", "AlarmLevel",
                    "Shift","LocationName","Address","City","State","PostalCode",
                    "Latitude","Longitude"]
    columns = [c for c in prefer_first if c in df.columns] + [c for c in df.columns if c not in prefer_first]
    cols3 = st.columns(3)
    for i, col in enumerate(columns):
        with cols3[i % 3]:
            current = defaults.get(col, None)
            key = f"incident_field_{col}"
            if col in lookups:
                options = lookups[col]
                idx = options.index(current) if isinstance(current, str) and current in options else None
                vals[col] = st.selectbox(col, options=options, index=idx, placeholder=f"Select {col}...", key=key)
            elif col in DATE_LIKE:
                d = pd.to_datetime(current).date() if pd.notna(current) else date.today()
                vals[col] = st.date_input(col, value=d, key=key)
            elif col in TIME_LIKE:
                vals[col] = st.text_input(col, value=str(current) if pd.notna(current) else "", placeholder="HH:MM", key=key)
            else:
                if col in df.select_dtypes(include="number").columns:
                    try:
                        base = float(current) if pd.notna(current) else 0.0
                    except Exception:
                        base = 0.0
                    vals[col] = st.number_input(col, value=base, key=key)
                else:
                    vals[col] = st.text_input(col, value=str(current) if pd.notna(current) else "", key=key)
    return vals

def upsert_row(df: pd.DataFrame, row: dict, key=PRIMARY_KEY) -> pd.DataFrame:
    if key not in df.columns:
        df[key] = pd.NA
    keys = df[key].astype(str) if not df.empty else pd.Series([], dtype=str)
    if str(row.get(key)) in keys.values:
        idx = df.index[keys == str(row[key])]
        for k, v in row.items():
            if k not in df.columns:
                df[k] = pd.NA
            df.loc[idx, k] = v
    else:
        for k in row.keys():
            if k not in df.columns:
                df[k] = pd.NA
        df = pd.concat([df, pd.DataFrame([row])], ignore_index=True)
    return df

def assignment_personnel_ui(data: Dict[str, pd.DataFrame], lookups: Dict[str, List[str]], inc_id: str):
    st.subheader("Assign Personnel")
    roster = ensure_columns(data.get("Personnel", pd.DataFrame()), PERSONNEL_SCHEMA)
    name_opts = roster["Name"].dropna().astype(str).tolist() if "Name" in roster.columns else []

    # Bulk add
    with st.expander("Add personnel from roster"):
        picked = st.multiselect("Select personnel", options=name_opts, key="assign_pick_people")
        roles = lookups.get("Role", ["OIC","Driver","Firefighter"])
        role = st.selectbox("Role (default for selected)", options=roles, index=0 if roles else None, key="assign_role_people")
        hours = st.number_input("Hours (each)", value=0.0, min_value=0.0, step=0.5, key="assign_hours_people")
        if st.button("Add selected personnel", key="btn_assign_people_add"):
            df = ensure_columns(data.get("Incident_Personnel", pd.DataFrame()), CHILD_TABLES["Incident_Personnel"])
            new = [{PRIMARY_KEY: inc_id, "Name": n, "Role": role, "Hours": hours} for n in picked]
            if new:
                data["Incident_Personnel"] = pd.concat([df, pd.DataFrame(new)], ignore_index=True)
                st.success(f"Added {len(new)} personnel.")

    # Current list + remove
    df_cur = ensure_columns(data.get("Incident_Personnel", pd.DataFrame()), CHILD_TABLES["Incident_Personnel"])
    view = df_cur[df_cur[PRIMARY_KEY].astype(str) == str(inc_id)].copy()
    if not view.empty:
        view = view.reset_index(drop=False).rename(columns={"index":"_row_index"})
        st.dataframe(view, use_container_width=True, hide_index=True)
        rm_indices = st.multiselect("Select rows to remove (by _row_index)", options=view["_row_index"].tolist(), key="assign_people_remove")
        if st.button("Remove selected personnel", key="btn_assign_people_remove"):
            df_cur = df_cur.drop(index=rm_indices, errors="ignore")
            data["Incident_Personnel"] = df_cur.reset_index(drop=True)
            st.success("Removed selected personnel.")

def assignment_apparatus_ui(data: Dict[str, pd.DataFrame], lookups: Dict[str, List[str]], inc_id: str):
    st.subheader("Assign Apparatus")
    roster = ensure_columns(data.get("Apparatus", pd.DataFrame()), APPARATUS_SCHEMA)
    label_col = "UnitNumber" if "UnitNumber" in roster.columns else ("CallSign" if "CallSign" in roster.columns else None)
    unit_opts = roster[label_col].dropna().astype(str).tolist() if label_col else []

    with st.expander("Add apparatus from roster"):
        picked = st.multiselect("Select apparatus", options=unit_opts, key="assign_pick_units")
        roles = ["Primary","Support","Water Supply","Staging"]
        role = st.selectbox("Role (default for selected units)", options=roles, index=0, key="assign_role_units")
        ut = st.selectbox("UnitType (optional)", options=lookups.get("UnitType", []), index=None, placeholder="Select...", key="assign_unittype_units")
        actions = st.multiselect("Actions (optional)", options=lookups.get("Action", []), key="assign_actions_units")
        if st.button("Add selected apparatus", key="btn_assign_units_add"):
            df = ensure_columns(data.get("Incident_Apparatus", pd.DataFrame()), CHILD_TABLES["Incident_Apparatus"])
            new = [{
                PRIMARY_KEY: inc_id,
                "Unit": u,
                "UnitType": ut,
                "Role": role,
                "Actions": "; ".join(actions) if actions else ""
            } for u in picked]
            if new:
                data["Incident_Apparatus"] = pd.concat([df, pd.DataFrame(new)], ignore_index=True)
                st.success(f"Added {len(new)} apparatus rows.")

    # Current list + remove
    df_cur = ensure_columns(data.get("Incident_Apparatus", pd.DataFrame()), CHILD_TABLES["Incident_Apparatus"])
    view = df_cur[df_cur[PRIMARY_KEY].astype(str) == str(inc_id)].copy()
    if not view.empty and PRIMARY_KEY in view.columns:
        view = view.reset_index(drop=False).rename(columns={"index":"_row_index"})
        st.dataframe(view, use_container_width=True, hide_index=True)
        rm_indices = st.multiselect("Select rows to remove (by _row_index)", options=view["_row_index"].tolist(), key="assign_units_remove")
        if st.button("Remove selected apparatus", key="btn_assign_units_remove"):
            df_cur = df_cur.drop(index=rm_indices, errors="ignore")
            data["Incident_Apparatus"] = df_cur.reset_index(drop=True)
            st.success("Removed selected apparatus.")

def analytics_reports(data: Dict[str, pd.DataFrame]):
    st.header("Reports")
    inc = data.get("Incidents", pd.DataFrame())
    if inc.empty:
        st.info("No incidents yet.")
        return

    if "IncidentDate" in inc.columns:
        inc["_date"] = pd.to_datetime(inc["IncidentDate"], errors="coerce")
        inc["_ym"] = inc["_date"].dt.to_period("M").astype(str)

    c1, c2 = st.columns(2)
    with c1:
        if "IncidentType" in inc.columns:
            st.subheader("Incidents by Type")
            by_type = inc["IncidentType"].value_counts().rename_axis("IncidentType").reset_index(name="Count")
            st.dataframe(by_type, use_container_width=True, hide_index=True)
    with c2:
        if "_ym" in inc.columns:
            st.subheader("Incidents by Month")
            by_month = inc["_ym"].value_counts().rename_axis("Month").reset_index(name="Count").sort_values("Month")
            st.dataframe(by_month, use_container_width=True, hide_index=True)

def incident_snapshot(data: Dict[str, pd.DataFrame], incident_number: str):
    per = ensure_columns(data.get("Incident_Personnel", pd.DataFrame()), CHILD_TABLES["Incident_Personnel"])
    app = ensure_columns(data.get("Incident_Apparatus", pd.DataFrame()), CHILD_TABLES["Incident_Apparatus"])
    per_view = per[per[PRIMARY_KEY].astype(str) == str(incident_number)]
    app_view = app[app[PRIMARY_KEY].astype(str) == str(incident_number)]

    total_personnel = len(per_view) if not per_view.empty else 0
    total_apparatus = len(app_view["Unit"].dropna()) if not app_view.empty and "Unit" in app_view.columns else 0

    st.subheader("Incident Snapshot")
    c1, c2 = st.columns(2)
    with c1:
        st.write(f"**Personnel on Scene:** {total_personnel}")
        if not per_view.empty:
            by_role = per_view["Role"].fillna("Unspecified").value_counts().rename_axis("Role").reset_index(name="Count")
            st.dataframe(by_role, use_container_width=True, hide_index=True)
    with c2:
        st.write(f"**Apparatus on Scene:** {total_apparatus}")
        if not app_view.empty:
            units = app_view["Unit"].dropna().astype(str).tolist() if "Unit" in app_view.columns else []
            st.write("**Units:** " + ", ".join(units))

def printable_incident(data: Dict[str, pd.DataFrame], incident_number: str):
    st.header(f"Incident Report — #{incident_number}")
    inc = data.get("Incidents", pd.DataFrame())
    row = inc[inc[PRIMARY_KEY].astype(str) == str(incident_number)]
    if row.empty:
        st.warning("Incident not found.")
        return
    rec = row.iloc[0].to_dict()

    cols = st.columns(2)
    left_keys = ["IncidentDate","IncidentTime","IncidentType","ResponsePriority","AlarmLevel","Shift"]
    right_keys = ["LocationName","Address","City","State","PostalCode"]
    with cols[0]:
        for k in left_keys:
            if k in rec:
                st.write(f"**{k}:** {rec.get(k,'')}")
    with cols[1]:
        for k in right_keys:
            if k in rec:
                st.write(f"**{k}:** {rec.get(k,'')}")

    per = ensure_columns(data.get("Incident_Personnel", pd.DataFrame()), CHILD_TABLES["Incident_Personnel"])
    app = ensure_columns(data.get("Incident_Apparatus", pd.DataFrame()), CHILD_TABLES["Incident_Apparatus"])
    per_view = per[per[PRIMARY_KEY].astype(str) == str(incident_number)]
    app_view = app[app[PRIMARY_KEY].astype(str) == str(incident_number)]

    st.subheader("On-Scene Summary")
    c1, c2 = st.columns(2)
    with c1:
        total_personnel = len(per_view) if not per_view.empty else 0
        st.write(f"**Personnel on Scene:** {total_personnel}")
        if not per_view.empty:
            by_role = per_view["Role"].fillna("Unspecified").value_counts().rename_axis("Role").reset_index(name="Count")
            st.dataframe(by_role, use_container_width=True, hide_index=True)
            roster = per_view.apply(lambda r: f"{r.get('Name','')} ({r.get('Role','')})", axis=1).tolist()
            st.write("**Roster:** " + ", ".join([x for x in roster if x and str(x).strip() != "()"]))
    with c2:
        total_apparatus = len(app_view["Unit"].dropna()) if not app_view.empty and "Unit" in app_view.columns else 0
        st.write(f"**Apparatus on Scene:** {total_apparatus}")
        if not app_view.empty:
            units = app_view["Unit"].dropna().astype(str).tolist() if "Unit" in app_view.columns else []
            st.write("**Units:** " + ", ".join(units))

    for t in ["Incident_Times","Incident_Personnel","Incident_Apparatus","Incident_Actions"]:
        df = data.get(t, pd.DataFrame())
        if not df.empty and PRIMARY_KEY in df.columns:
            view = df[df[PRIMARY_KEY].astype(str) == str(incident_number)]
            if not view.empty:
                st.subheader(t.replace("_"," "))
                st.dataframe(view, use_container_width=True, hide_index=True)

# ---------- App ----------
st.sidebar.title("🚒 Fire Incident DB")
file_path = st.sidebar.text_input("Excel path", value=DEFAULT_FILE, key="sidebar_path")
uploaded = st.sidebar.file_uploader("Upload/replace workbook (.xlsx)", type=["xlsx"], key="sidebar_upload")
if uploaded:
    with open(file_path, "wb") as f:
        f.write(uploaded.read())
    st.sidebar.success(f"Saved to {file_path}")

data = {}
if os.path.exists(file_path):
    data = load_workbook(file_path)

if not data:
    st.info("Upload or point to your Excel workbook to begin.")
    st.stop()

# Ensure core sheets exist
if "Incidents" not in data:
    data["Incidents"] = pd.DataFrame(columns=[PRIMARY_KEY])
for t, cols in CHILD_TABLES.items():
    if t not in data:
        data[t] = pd.DataFrame(columns=cols)

# Ensure rosters exist
if "Personnel" not in data:
    data["Personnel"] = pd.DataFrame(columns=PERSONNEL_SCHEMA)
else:
    data["Personnel"] = ensure_columns(data["Personnel"], PERSONNEL_SCHEMA)
if "Apparatus" not in data:
    data["Apparatus"] = pd.DataFrame(columns=APPARATUS_SCHEMA)
else:
    data["Apparatus"] = ensure_columns(data["Apparatus"], APPARATUS_SCHEMA)

lookups = get_lookups(data)

tab1, tab2, tab3, tab4, tab5, tab6, tab7 = st.tabs(["Browse","Add / Edit","Assign to Incident","Related (Raw)","Rosters","Reports","Export"])

with tab1:
    st.header("Browse & Filter Incidents")
    base = data["Incidents"].copy()
    fc1, fc2, fc3, fc4 = st.columns(4)
    with fc1:
        val_type = st.selectbox("IncidentType", options=[""]+lookups.get("IncidentType", []), key="filter_type")
        if val_type:
            base = base[base.get("IncidentType","").astype(str) == val_type]
    with fc2:
        val_prio = st.selectbox("ResponsePriority", options=[""]+lookups.get("ResponsePriority", []), key="filter_prio")
        if val_prio:
            base = base[base.get("ResponsePriority","").astype(str) == val_prio]
    with fc3:
        city = st.text_input("City contains", key="filter_city")
        if city:
            base = base[base.get("City","").astype(str).str.contains(city, case=False, na=False)]
    with fc4:
        dr = st.date_input("Date range", [], key="filter_dates")
        if isinstance(dr, list) and len(dr) == 2:
            start, end = pd.to_datetime(dr[0]), pd.to_datetime(dr[1])
            base = base[(pd.to_datetime(base.get("IncidentDate", pd.NaT), errors="coerce") >= start) &
                        (pd.to_datetime(base.get("IncidentDate", pd.NaT), errors="coerce") <= end)]
    st.dataframe(base, use_container_width=True, hide_index=True)

with tab2:
    st.header("Add / Edit Incident")
    master = data["Incidents"]
    mode = st.radio("Mode", ["Add","Edit"], horizontal=True, key="mode_incident")
    defaults = {}
    selected = None
    if mode == "Edit" and not master.empty and PRIMARY_KEY in master.columns:
        options = master[PRIMARY_KEY].dropna().astype(str).tolist()
        selected = st.selectbox("Select IncidentNumber", options=options, index=None, placeholder="Choose...", key="pick_incident_edit")
        if selected:
            defaults = master[master[PRIMARY_KEY].astype(str) == selected].iloc[0].to_dict()

    vals = render_dynamic_form(master, lookups, defaults)
    if "IncidentDate" in vals:
        vals["IncidentDate"] = pd.to_datetime(vals["IncidentDate"])
    if st.button("Save Incident", key="btn_save_incident"):
        data["Incidents"] = upsert_row(master, vals, key=PRIMARY_KEY)
        st.success("Saved.")
    if mode == "Edit" and selected:
        with st.expander("Incident Snapshot", expanded=True):
            incident_snapshot(data, selected)

with tab3:
    st.header("Assign to Incident")
    if data["Incidents"].empty:
        st.info("Add an incident first.")
    else:
        inc_id = st.selectbox("IncidentNumber", options=data["Incidents"][PRIMARY_KEY].dropna().astype(str).tolist(), index=None, key="pick_incident_assign")
        if inc_id:
            c1, c2 = st.columns(2)
            with c1:
                assignment_personnel_ui(data, lookups, inc_id)
            with c2:
                assignment_apparatus_ui(data, lookups, inc_id)

with tab4:
    st.header("Related (Raw tables)")
    st.write("For power users: view & append directly to the raw related tables.")
    if data["Incidents"].empty:
        st.info("Add an incident first.")
    else:
        inc_id = st.selectbox("IncidentNumber", options=data["Incidents"][PRIMARY_KEY].dropna().astype(str).tolist(), index=None, key="pick_incident_related")
        if inc_id:
            for t in ["Incident_Times","Incident_Personnel","Incident_Apparatus","Incident_Actions"]:
                st.subheader(t)
                df = ensure_columns(data.get(t, pd.DataFrame()), CHILD_TABLES[t])
                st.dataframe(df[df[PRIMARY_KEY].astype(str) == str(inc_id)], use_container_width=True, hide_index=True)

with tab5:
    st.header("Rosters (Master Lists)")
    st.write("Edit your master Personnel & Apparatus lists here. These are used for assignment.")
    # Personnel Roster editor
    personnel = ensure_columns(data.get("Personnel", pd.DataFrame()), PERSONNEL_SCHEMA)
    personnel_edit = st.data_editor(personnel, num_rows="dynamic", use_container_width=True, key="editor_personnel_roster")
    if st.button("Save Personnel Roster", key="save_personnel_roster"):
        data["Personnel"] = personnel_edit
        st.success("Personnel roster saved (in memory). Use Export to write to Excel.")
    # Apparatus Roster editor
    apparatus = ensure_columns(data.get("Apparatus", pd.DataFrame()), APPARATUS_SCHEMA)
    apparatus_edit = st.data_editor(apparatus, num_rows="dynamic", use_container_width=True, key="editor_apparatus_roster")
    if st.button("Save Apparatus Roster", key="save_apparatus_roster"):
        data["Apparatus"] = apparatus_edit
        st.success("Apparatus roster saved (in memory). Use Export to write to Excel.")

with tab6:
    analytics_reports(data)

with tab7:
    st.header("Export")
    if st.button("Build Excel for Download", key="btn_build_export"):
        payload = save_workbook_to_bytes(data)
        st.download_button("Download Excel", data=payload, file_name="fire_incident_db_export.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key="download_export")
    if st.button("Overwrite Source File", key="btn_overwrite_source"):
        payload = save_workbook_to_bytes(data)
        with open(file_path, "wb") as f:
            f.write(payload)
        st.success(f"Overwrote: {file_path}")
