
import os, io
from datetime import datetime, date
from typing import Dict, List
import pandas as pd
import streamlit as st

st.set_page_config(page_title="Fire Incident Reports", page_icon="📝", layout="wide")

DEFAULT_FILE = os.path.join(os.path.dirname(__file__), "fire_incident_db.xlsx")
PRIMARY_KEY = "IncidentNumber"
CHILD_TABLES = {
    "Incident_Times": ["IncidentNumber","Alarm","Enroute","Arrival","Clear"],
    "Incident_Personnel": ["IncidentNumber","Name","Role","Hours"],
    "Incident_Apparatus": ["IncidentNumber","Unit","UnitType","Role","Actions"],
    "Incident_Actions": ["IncidentNumber","Action","Notes"],
}
PERSONNEL_SCHEMA = ["PersonnelID","Name","UnitNumber","Rank","Badge","Phone","Email","Address","City","State","PostalCode","Certifications","Active","FirstName","LastName","FullName"]
APPARATUS_SCHEMA = ["ApparatusID","UnitNumber","CallSign","UnitType","GPM","TankSize","SeatingCapacity","Station","Active","Name"]
USERS_SCHEMA = ["Username","Password","Role","FullName","Active"]

LOOKUP_SHEETS = {
    "List_IncidentType": "IncidentType",
    "List_AlarmLevel": "AlarmLevel",
    "List_ResponsePriority": "ResponsePriority",
    "List_PersonnelRoles": "Role",
    "List_UnitTypes": "UnitType",
    "List_Actions": "Action",
    "List_States": "State",
}

# ---------------- Utilities ----------------
def load_workbook(path: str) -> Dict[str, pd.DataFrame]:
    try:
        xls = pd.ExcelFile(path)
        return {name: xls.parse(name) for name in xls.sheet_names}
    except Exception as e:
        st.error(f"Failed to load workbook: {e}")
        return {}

def save_workbook_to_bytes(dfs: Dict[str, pd.DataFrame]) -> bytes:
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="xlsxwriter") as writer:
        for sheet, df in dfs.items():
            df.to_excel(writer, sheet_name=sheet, index=False)
    buf.seek(0)
    return buf.read()

def save_to_path(dfs: Dict[str, pd.DataFrame], path: str):
    try:
        with pd.ExcelWriter(path, engine="xlsxwriter") as writer:
            for sheet, df in dfs.items():
                df.to_excel(writer, sheet_name=sheet, index=False)
        return True, None
    except Exception as e:
        return False, str(e)

def ensure_columns(df: pd.DataFrame, cols: List[str]) -> pd.DataFrame:
    if df is None:
        df = pd.DataFrame()
    for c in cols:
        if c not in df.columns:
            df[c] = pd.NA
    return df

def ensure_table(data: Dict[str, pd.DataFrame], name: str, cols: List[str]):
    data[name] = ensure_columns(data.get(name, pd.DataFrame()), cols)

def get_lookups(data: Dict[str, pd.DataFrame]) -> Dict[str, List[str]]:
    out: Dict[str, List[str]] = {}
    for sheet, col in LOOKUP_SHEETS.items():
        if sheet in data and not data[sheet].empty:
            header = data[sheet].columns[0]
            out[col] = data[sheet][header].dropna().astype(str).tolist()
    return out

def upsert_row(df: pd.DataFrame, row: dict, key=PRIMARY_KEY) -> pd.DataFrame:
    df = ensure_columns(df, list(row.keys()) + [key])
    if key not in df.columns:
        df[key] = pd.NA
    keys = df[key].astype(str) if not df.empty else pd.Series([], dtype=str)
    if str(row.get(key)) in keys.values:
        idx = df.index[keys == str(row[key])]
        for k, v in row.items():
            if k not in df.columns: df[k] = pd.NA
            df.loc[idx, k] = v
    else:
        for k in row.keys():
            if k not in df.columns: df[k] = pd.NA
        df = pd.concat([df, pd.DataFrame([row])], ignore_index=True)
    return df

def name_rank_first_last(row: pd.Series) -> str:
    fn = str(row.get("FirstName") or "").strip()
    ln = str(row.get("LastName") or "").strip()
    rk = str(row.get("Rank") or "").strip()
    parts = [p for p in [rk, fn, ln] if p]
    return " ".join(parts).strip()

def build_person_options(df: pd.DataFrame) -> list:
    # prefer Name/FullName, else synthesize
    if "Name" in df and df["Name"].notna().any():
        s = df["Name"].astype(str)
    elif "FullName" in df and df["FullName"].notna().any():
        s = df["FullName"].astype(str)
    elif all(c in df.columns for c in ["FirstName","LastName","Rank"]):
        s = df.apply(name_rank_first_last, axis=1)
    elif all(c in df.columns for c in ["FirstName","LastName"]):
        s = (df["FirstName"].fillna("").astype(str).str.strip() + " " + df["LastName"].fillna("").astype(str).str.strip()).str.strip()
    else:
        s = pd.Series([], dtype=str)
    # no Active filter by default to avoid hiding data unintentionally
    opts = s.dropna().map(lambda x: x.strip()).replace("", pd.NA).dropna().unique().tolist()
    return sorted(set(opts))

def build_unit_options(df: pd.DataFrame) -> list:
    for col in ["UnitNumber","CallSign","Name"]:
        if col in df.columns and df[col].notna().any():
            s = df[col].astype(str); break
    else:
        s = pd.Series([], dtype=str)
    opts = s.dropna().map(lambda x: x.strip()).replace("", pd.NA).dropna().unique().tolist()
    return sorted(set(opts))

def repair_rosters(data: Dict[str, pd.DataFrame]) -> Dict[str, pd.DataFrame]:
    # Personnel
    p = ensure_columns(data.get("Personnel", pd.DataFrame()), PERSONNEL_SCHEMA).copy()
    if not p.empty:
        # fill Name/FullName if blank
        mask_name_blank = p["Name"].isna() | (p["Name"].astype(str).str.strip()=="")
        p.loc[mask_name_blank, "Name"] = p.loc[mask_name_blank].apply(name_rank_first_last, axis=1)
        mask_full_blank = p["FullName"].isna() | (p["FullName"].astype(str).str.strip()=="")
        p.loc[mask_full_blank, "FullName"] = p.loc[mask_full_blank].apply(name_rank_first_last, axis=1)
        # default Active
        if "Active" in p.columns:
            m = p["Active"].isna() | (p["Active"].astype(str).strip()=="")
            p.loc[m, "Active"] = "Yes"
        else:
            p["Active"] = "Yes"
    data["Personnel"] = p
    # Apparatus
    a = ensure_columns(data.get("Apparatus", pd.DataFrame()), APPARATUS_SCHEMA).copy()
    if not a.empty:
        if "Active" in a.columns:
            m = a["Active"].isna() | (a["Active"].astype(str).strip()=="")
            a.loc[m, "Active"] = "Yes"
        else:
            a["Active"] = "Yes"
    data["Apparatus"] = a
    return data

# ---------------- App Boot ----------------
st.sidebar.title("📝 Fire Incident Reports — v5.0 Clean Reset")
file_path = st.sidebar.text_input("Excel path", value=DEFAULT_FILE, key="path_input_v5")
uploaded = st.sidebar.file_uploader("Upload/replace workbook (.xlsx)", type=["xlsx"], key="upload_v5")
if uploaded:
    with open(file_path, "wb") as f: f.write(uploaded.read())
    st.sidebar.success(f"Saved to {file_path}")
st.session_state.setdefault("autosave", True)
st.session_state["autosave"] = st.sidebar.toggle("Autosave to Excel", value=True, key="autosave_v5")

data: Dict[str, pd.DataFrame] = {}
if os.path.exists(file_path):
    data = load_workbook(file_path)
else:
    st.info("Upload or point to your Excel workbook to begin.")
    st.stop()

# Ensure required tables
ensure_table(data, "Incidents", [
    PRIMARY_KEY,"IncidentDate","IncidentTime","IncidentType","ResponsePriority","AlarmLevel","Shift",
    "LocationName","Address","City","State","PostalCode","Latitude","Longitude",
    "Narrative","Status","CreatedBy","SubmittedAt","ReviewedBy","ReviewedAt","ReviewerComments"
])
ensure_table(data, "Personnel", PERSONNEL_SCHEMA)
ensure_table(data, "Apparatus", APPARATUS_SCHEMA)
for t, cols in CHILD_TABLES.items(): ensure_table(data, t, cols)

# Auto-repair once on load (non-destructive)
data = repair_rosters(data)

lookups = get_lookups(data)

# ---- Tabs
tabs = st.tabs(["Write Report","Review Queue","Rosters","Print","Export"])

# ---- Write Report
with tabs[0]:
    st.header("Write Report")
    master = data["Incidents"]
    mode = st.radio("Mode", ["New","Edit"], horizontal=True, key="mode_write_v5")
    defaults = {}; selected = None
    if mode == "Edit" and not master.empty:
        options = master[PRIMARY_KEY].dropna().astype(str).tolist() if PRIMARY_KEY in master.columns else []
        selected = st.selectbox("Select IncidentNumber", options=options, index=None, placeholder="Choose...", key="pick_edit_write_v5")
        if selected:
            defaults = master[master[PRIMARY_KEY].astype(str) == selected].iloc[0].to_dict()

    with st.container(border=True):
        st.subheader("Incident Details")
        c1, c2, c3 = st.columns(3)
        inc_num = c1.text_input("IncidentNumber", value=str(defaults.get(PRIMARY_KEY,"")) if defaults else "", key="w_inc_num_v5")
        inc_date = c2.date_input("IncidentDate", value=pd.to_datetime(defaults.get("IncidentDate")).date() if defaults.get("IncidentDate") is not None and str(defaults.get("IncidentDate")) != "NaT" else date.today(), key="w_inc_date_v5")
        inc_time = c3.text_input("IncidentTime (HH:MM)", value=str(defaults.get("IncidentTime","")) if defaults else "", key="w_inc_time_v5")
        c4, c5, c6 = st.columns(3)
        inc_type = c4.selectbox("IncidentType", options=[""]+lookups.get("IncidentType", []), index=0, key="w_type_v5")
        inc_prio = c5.selectbox("ResponsePriority", options=[""]+lookups.get("ResponsePriority", []), index=0, key="w_prio_v5")
        inc_alarm = c6.selectbox("AlarmLevel", options=[""]+lookups.get("AlarmLevel", []), index=0, key="w_alarm_v5")
        c7, c8, c9 = st.columns(3)
        loc_name = c7.text_input("LocationName", value=str(defaults.get("LocationName","")) if defaults else "", key="w_locname_v5")
        addr = c8.text_input("Address", value=str(defaults.get("Address","")) if defaults else "", key="w_addr_v5")
        city = c9.text_input("City", value=str(defaults.get("City","")) if defaults else "", key="w_city_v5")
        c10, c11, c12 = st.columns(3)
        state = c10.text_input("State", value=str(defaults.get("State","")) if defaults else "", key="w_state_v5")
        postal = c11.text_input("PostalCode", value=str(defaults.get("PostalCode","")) if defaults else "", key="w_postal_v5")
        shift = c12.text_input("Shift", value=str(defaults.get("Shift","")) if defaults else "", key="w_shift_v5")

    with st.container(border=True):
        st.subheader("Narrative")
        narrative = st.text_area("Write full narrative here", value=str(defaults.get("Narrative","")) if defaults else "", height=320, key="w_narrative_v5")

    with st.container(border=True):
        st.subheader("All Members on Scene")
        people_df = ensure_columns(data.get("Personnel", pd.DataFrame()), PERSONNEL_SCHEMA)
        person_opts = build_person_options(people_df)
        picked_people = st.multiselect("Pick members", options=person_opts, key="w_pick_people_v5")
        roles = lookups.get("Role", ["OIC","Driver","Firefighter"])
        cc = st.columns(3)
        role_default = cc[0].selectbox("Default Role", options=roles, index=0 if roles else None, key="w_role_default_v5")
        hours_default = cc[1].number_input("Default Hours", value=0.0, min_value=0.0, step=0.5, key="w_hours_default_v5")
        if cc[2].button("Add Selected Members", key="w_add_people_btn_v5"):
            if inc_num:
                df = ensure_columns(data.get("Incident_Personnel", pd.DataFrame()), CHILD_TABLES["Incident_Personnel"])
                new = [{PRIMARY_KEY: inc_num, "Name": n, "Role": role_default, "Hours": hours_default} for n in picked_people]
                if new:
                    data["Incident_Personnel"] = pd.concat([df, pd.DataFrame(new)], ignore_index=True)
                    if st.session_state.get("autosave", True):
                        save_to_path(data, file_path)
                    st.success(f"Added {len(new)} member(s).")
        cur_per = ensure_columns(data.get("Incident_Personnel", pd.DataFrame()), CHILD_TABLES["Incident_Personnel"])
        cur_view = cur_per[cur_per[PRIMARY_KEY].astype(str) == str(inc_num)]
        st.write(f"**Total Personnel on Scene:** {len(cur_view) if not cur_view.empty else 0}")
        st.dataframe(cur_view, use_container_width=True, hide_index=True)

    with st.container(border=True):
        st.subheader("Apparatus on Scene")
        app_df = ensure_columns(data.get("Apparatus", pd.DataFrame()), APPARATUS_SCHEMA)
        unit_opts = build_unit_options(app_df)
        picked_units = st.multiselect("Pick apparatus units", options=unit_opts, key="w_pick_units_v5")
        cc2 = st.columns(3)
        unit_type = cc2[0].selectbox("UnitType", options=[""]+lookups.get("UnitType", []), index=0, key="w_unit_type_v5")
        unit_role = cc2[1].selectbox("Role", options=["Primary","Support","Water Supply","Staging"], index=0, key="w_unit_role_v5")
        if cc2[2].button("Add Selected Units", key="w_add_units_btn_v5"):
            if inc_num:
                df = ensure_columns(data.get("Incident_Apparatus", pd.DataFrame()), CHILD_TABLES["Incident_Apparatus"])
                new = [{PRIMARY_KEY: inc_num, "Unit": u, "UnitType": (unit_type if unit_type else None), "Role": unit_role, "Actions": ""} for u in picked_units]
                if new:
                    data["Incident_Apparatus"] = pd.concat([df, pd.DataFrame(new)], ignore_index=True)
                    if st.session_state.get("autosave", True):
                        save_to_path(data, file_path)
                    st.success(f"Added {len(new)} unit(s).")
        cur_app = ensure_columns(data.get("Incident_Apparatus", pd.DataFrame()), CHILD_TABLES["Incident_Apparatus"])
        cur_app_view = cur_app[cur_app[PRIMARY_KEY].astype(str) == str(inc_num)]
        st.write(f"**Total Apparatus on Scene:** {len(cur_app_view) if not cur_app_view.empty else 0}")
        st.dataframe(cur_app_view, use_container_width=True, hide_index=True)

    with st.container(border=True):
        st.subheader("Times (optional)")
        t1, t2, t3, t4 = st.columns(4)
        alarm = t1.text_input("Alarm (HH:MM)", key="w_alarm_time_v5")
        enroute = t2.text_input("Enroute (HH:MM)", key="w_enroute_time_v5")
        arrival = t3.text_input("Arrival (HH:MM)", key="w_arrival_time_v5")
        clear = t4.text_input("Clear (HH:MM)", key="w_clear_time_v5")
        if st.button("Save Times", key="w_save_times_v5"):
            times = data["Incident_Times"]
            new = {PRIMARY_KEY: inc_num, "Alarm": alarm, "Enroute": enroute, "Arrival": arrival, "Clear": clear}
            data["Incident_Times"] = upsert_row(times, new, key=PRIMARY_KEY)
            if st.session_state.get("autosave", True):
                save_to_path(data, file_path)
            st.success("Times saved.")

    row_vals = {
        PRIMARY_KEY: inc_num,
        "IncidentDate": pd.to_datetime(inc_date),
        "IncidentTime": inc_time,
        "IncidentType": inc_type,
        "ResponsePriority": inc_prio,
        "AlarmLevel": inc_alarm,
        "LocationName": loc_name,
        "Address": addr,
        "City": city,
        "State": state,
        "PostalCode": postal,
        "Shift": shift,
        "Narrative": narrative,
        "CreatedBy": "member",  # keep simple; hook up auth later if needed
    }
    a = st.columns(3)
    if a[0].button("Save Draft", key="w_save_draft_v5"):
        row_vals["Status"] = "Draft"
        data["Incidents"] = upsert_row(data["Incidents"], row_vals, key=PRIMARY_KEY)
        if st.session_state.get("autosave", True):
            save_to_path(data, file_path)
        st.success("Draft saved.")
    if a[1].button("Submit for Review", key="w_submit_review_v5"):
        row_vals["Status"] = "Submitted"; row_vals["SubmittedAt"] = datetime.now().strftime("%Y-%m-%d %H:%M")
        data["Incidents"] = upsert_row(data["Incidents"], row_vals, key=PRIMARY_KEY)
        if st.session_state.get("autosave", True):
            save_to_path(data, file_path)
        st.success("Submitted for review.")

# ---- Review Queue
with tabs[1]:
    st.header("Review Queue")
    pending = data["Incidents"][data["Incidents"]["Status"].astype(str) == "Submitted"]
    st.dataframe(pending, use_container_width=True, hide_index=True, key="grid_pending_v5")
    sel = None
    if not pending.empty:
        sel = st.selectbox("Pick an Incident to review", options=pending[PRIMARY_KEY].astype(str).tolist(), index=None, placeholder="Choose...", key="pick_review_queue_v5")
    if sel:
        rec = data["Incidents"][data["Incidents"][PRIMARY_KEY].astype(str) == sel].iloc[0].to_dict()
        st.json(rec, expanded=False)
        comments = st.text_area("Reviewer Comments", key="rev_comments_queue_v5")
        c = st.columns(3)
        if c[0].button("Approve", key="btn_approve_queue_v5"):
            row = rec; row["Status"] = "Approved"; row["ReviewedBy"] = "reviewer"; row["ReviewedAt"] = datetime.now().strftime("%Y-%m-%d %H:%M"); row["ReviewerComments"] = comments
            data["Incidents"] = upsert_row(data["Incidents"], row, key=PRIMARY_KEY)
            if st.session_state.get("autosave", True):
                save_to_path(data, file_path)
            st.success("Approved.")
        if c[1].button("Reject", key="btn_reject_queue_v5"):
            row = rec; row["Status"] = "Rejected"; row["ReviewedBy"] = "reviewer"; row["ReviewedAt"] = datetime.now().strftime("%Y-%m-%d %H:%M"); row["ReviewerComments"] = comments or "Please revise."
            data["Incidents"] = upsert_row(data["Incidents"], row, key=PRIMARY_KEY)
            if st.session_state.get("autosave", True):
                save_to_path(data, file_path)
            st.warning("Rejected.")
        if c[2].button("Send back to Draft", key="btn_backtodraft_queue_v5"):
            row = rec; row["Status"] = "Draft"; row["ReviewerComments"] = comments
            data["Incidents"] = upsert_row(data["Incidents"], row, key=PRIMARY_KEY)
            if st.session_state.get("autosave", True):
                save_to_path(data, file_path)
            st.info("Moved back to Draft.")

# ---- Rosters
with tabs[2]:
    st.header("Rosters")
    st.caption("Edit, then click Save to write changes into the Excel workbook. Use 'Repair roster now' if names are blank.")
    personnel = ensure_columns(data.get("Personnel", pd.DataFrame()), PERSONNEL_SCHEMA)
    personnel_edit = st.data_editor(personnel, num_rows="dynamic", use_container_width=True, key="editor_personnel_v5")
    apparatus = ensure_columns(data.get("Apparatus", pd.DataFrame()), APPARATUS_SCHEMA)
    apparatus_edit = st.data_editor(apparatus, num_rows="dynamic", use_container_width=True, key="editor_apparatus_v5")
    c = st.columns(3)
    if c[0].button("Save Personnel to Excel", key="save_personnel_v5"):
        data["Personnel"] = ensure_columns(personnel_edit, PERSONNEL_SCHEMA)
        ok, err = save_to_path(data, file_path)
        st.success("Saved.") if ok else st.error(err)
    if c[1].button("Save Apparatus to Excel", key="save_app_v5"):
        data["Apparatus"] = ensure_columns(apparatus_edit, APPARATUS_SCHEMA)
        ok, err = save_to_path(data, file_path)
        st.success("Saved.") if ok else st.error(err)
    if c[2].button("Repair roster now (fill names & Active)", key="repair_now_v5"):
        data = repair_rosters(data)
        ok, err = save_to_path(data, file_path)
        if ok: st.success("Roster repaired and saved. Reopen this tab or reload to see updates.")
        else: st.error(err)

# ---- Print
with tabs[3]:
    st.header("Print")
    status = st.selectbox("Filter by Status", options=["","Approved","Submitted","Draft","Rejected"], key="print_status_v5")
    base = data["Incidents"].copy()
    if status: base = base[base["Status"].astype(str) == status]
    st.dataframe(base, use_container_width=True, hide_index=True, key="grid_print_v5")
    sel = None
    if not base.empty:
        sel = st.selectbox("Pick an Incident", options=base[PRIMARY_KEY].astype(str).tolist(), index=None, key="print_pick_v5")
    if sel:
        rec = data["Incidents"][data["Incidents"][PRIMARY_KEY].astype(str) == sel].iloc[0].to_dict()
        st.json(rec, expanded=False)

# ---- Export
with tabs[4]:
    st.header("Export")
    if st.button("Build Excel for Download", key="btn_build_export_v5"):
        payload = save_workbook_to_bytes(data)
        st.download_button("Download Excel", data=payload, file_name="fire_incident_db_export.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key="download_export_v5")
    if st.button("Overwrite Source File Now", key="btn_overwrite_source_v5"):
        ok, err = save_to_path(data, file_path)
        if ok: st.success(f"Wrote: {file_path}")
        else: st.error(f"Failed: {err}")
