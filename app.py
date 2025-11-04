
import os
import io
from datetime import datetime, date
from typing import Dict, List

import pandas as pd
import streamlit as st

st.set_page_config(page_title="Fire Incident Reports", page_icon="📝", layout="wide")

# ================= Settings / Schemas =================
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
USERS_SCHEMA = ["Username","Password","Role","FullName","Active"]  # Demo only

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

# ================= Utils =================
def load_workbook(path: str) -> Dict[str, pd.DataFrame]:
    try:
        xls = pd.ExcelFile(path)
        return {name: xls.parse(name) for name in xls.sheet_names}
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

def autosave_if_enabled(data, file_path):
    if st.session_state.get("autosave", True):
        ok, err = save_to_path(data, file_path)
        if not ok:
            st.error(f"Autosave failed: {err}")

def analytics_reports(data: Dict[str, pd.DataFrame]):
    st.header("Reports")
    inc = data.get("Incidents", pd.DataFrame())
    if inc.empty:
        st.info("No incidents yet."); return
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
    c1, c2 = st.columns(2)
    with c1:
        st.write(f"**Personnel on Scene:** {total_personnel}")
        if not per_view.empty:
            by_role = per_view["Role"].fillna("Unspecified").value_counts().rename_axis("Role").reset_index(name="Count")
            st.dataframe(by_role, use_container_width=True, hide_index=True)
            roster = per_view.apply(lambda r: f"{r.get('Name','')} ({r.get('Role','')})", axis=1).tolist()
            st.write("**Roster:** " + ", ".join([x for x in roster if x and str(x).strip() != "()"]))
    with c2:
        st.write(f"**Apparatus on Scene:** {total_apparatus}")
        if not app_view.empty:
            units = app_view["Unit"].dropna().astype(str).tolist() if "Unit" in app_view.columns else []
            st.write("**Units:** " + ", ".join(units))

def printable_incident(data: Dict[str, pd.DataFrame], incident_number: str):
    st.header(f"Incident Report — #{incident_number}")
    inc = data.get("Incidents", pd.DataFrame())
    row = inc[inc[PRIMARY_KEY].astype(str) == str(incident_number)]
    if row.empty: st.warning("Incident not found."); return
    rec = row.iloc[0].to_dict()
    cols = st.columns(2)
    left_keys = ["IncidentNumber","IncidentDate","IncidentTime","IncidentType","ResponsePriority","AlarmLevel","Shift"]
    right_keys = ["LocationName","Address","City","State","PostalCode"]
    with cols[0]:
        for k in left_keys:
            if k in rec: st.write(f"**{k}:** {rec.get(k,'')}")
    with cols[1]:
        for k in right_keys:
            if k in rec: st.write(f"**{k}:** {rec.get(k,'')}")
    st.write(f"**Status:** {rec.get('Status','')}  —  **CreatedBy:** {rec.get('CreatedBy','')}")
    if rec.get("ReviewerComments"): st.write(f"**ReviewerComments:** {rec.get('ReviewerComments')}")
    if "Narrative" in rec and pd.notna(rec["Narrative"]):
        st.subheader("Narrative")
        st.write(rec["Narrative"])
    incident_snapshot(data, str(rec.get('IncidentNumber','')))
    for t in ["Incident_Times","Incident_Personnel","Incident_Apparatus","Incident_Actions"]:
        df = data.get(t, pd.DataFrame())
        if not df.empty and PRIMARY_KEY in df.columns:
            view = df[df[PRIMARY_KEY].astype(str) == str(rec.get('IncidentNumber',''))]
            if not view.empty:
                st.subheader(t.replace("_"," ")); st.dataframe(view, use_container_width=True, hide_index=True)

# ================= Auth =================
def ensure_default_users(data: Dict[str, pd.DataFrame]):
    users = ensure_columns(data.get("Users", pd.DataFrame()), USERS_SCHEMA)
    if users.empty:
        users = pd.DataFrame([
            {"Username":"admin","Password":"admin","Role":"Admin","FullName":"Administrator","Active":"Yes"},
            {"Username":"review","Password":"review","Role":"Reviewer","FullName":"Reviewer","Active":"Yes"},
            {"Username":"member","Password":"member","Role":"Member","FullName":"Member User","Active":"Yes"},
        ])
    data["Users"] = users; return data

def sign_in_ui(users_df: pd.DataFrame):
    st.header("Sign In")
    u = st.text_input("Username", key="login_user")
    p = st.text_input("Password", type="password", key="login_pass")
    ok = st.button("Sign In", key="btn_login")
    if ok:
        row = users_df[(users_df["Username"].astype(str)==u) & (users_df["Password"].astype(str)==p) & (users_df["Active"].astype(str).str.lower()=="yes")]
        if not row.empty:
            st.session_state["user"] = {"username": u, "role": row.iloc[0]["Role"], "name": row.iloc[0].get("FullName", u)}
            st.success(f"Welcome, {st.session_state['user']['name']}!"); st.experimental_rerun()
        else:
            st.error("Invalid credentials or inactive user.")

def sign_out_button():
    if st.button("Sign Out", key="btn_logout"):
        st.session_state.pop("user", None); st.experimental_rerun()

# ================= Assignment helpers =================
def bulk_add_personnel(data: Dict[str, pd.DataFrame], inc_id: str, names: List[str], role: str, hours: float):
    df = ensure_columns(data.get("Incident_Personnel", pd.DataFrame()), CHILD_TABLES["Incident_Personnel"])
    new = [{PRIMARY_KEY: inc_id, "Name": n, "Role": role, "Hours": hours} for n in names]
    if new:
        data["Incident_Personnel"] = pd.concat([df, pd.DataFrame(new)], ignore_index=True)

def bulk_add_units(data: Dict[str, pd.DataFrame], inc_id: str, units: List[str], role: str, unit_type: str, actions: List[str]):
    df = ensure_columns(data.get("Incident_Apparatus", pd.DataFrame()), CHILD_TABLES["Incident_Apparatus"])
    new = [{
        PRIMARY_KEY: inc_id,
        "Unit": u,
        "UnitType": unit_type if unit_type else None,
        "Role": role,
        "Actions": "; ".join(actions) if actions else ""
    } for u in units]
    if new:
        data["Incident_Apparatus"] = pd.concat([df, pd.DataFrame(new)], ignore_index=True)

# ================= App Boot =================
st.sidebar.title("📝 Fire Incident Reports")
file_path = st.sidebar.text_input("Excel path", value=DEFAULT_FILE, key="sidebar_path")
uploaded = st.sidebar.file_uploader("Upload/replace workbook (.xlsx)", type=["xlsx"], key="sidebar_upload")
if uploaded:
    with open(file_path, "wb") as f: f.write(uploaded.read())
    st.sidebar.success(f"Saved to {file_path}")
st.session_state.setdefault("autosave", True)
st.session_state["autosave"] = st.sidebar.toggle("Autosave to Excel", value=True, help="Write changes to Excel immediately")

data: Dict[str, pd.DataFrame] = {}
if os.path.exists(file_path): data = load_workbook(file_path)
if not data: st.info("Upload or point to your Excel workbook to begin."); st.stop()

# Ensure tables exist (Narrative added)
ensure_table(data, "Incidents", [
    PRIMARY_KEY,"IncidentDate","IncidentTime","IncidentType","ResponsePriority","AlarmLevel","Shift",
    "LocationName","Address","City","State","PostalCode","Latitude","Longitude",
    "Narrative",
    "Status","CreatedBy","SubmittedAt","ReviewedBy","ReviewedAt","ReviewerComments"
])
for t, cols in CHILD_TABLES.items(): ensure_table(data, t, cols)
ensure_table(data, "Personnel", PERSONNEL_SCHEMA)
ensure_table(data, "Apparatus", APPARATUS_SCHEMA)
data = ensure_default_users(data)
lookups = get_lookups(data)

# Auth
users_df = data["Users"]
if "user" not in st.session_state:
    sign_in_ui(users_df); st.stop()
user = st.session_state["user"]
st.sidebar.write(f"**Logged in as:** {user['name']}  \\nRole: {user['role']}")
sign_out_button()

# ================= Tabs =================
tabs = st.tabs([
    "Write Report",
    "Review Queue",
    "Browse",
    "Rosters",
    "Reports",
    "Print",
    "Export",
])

# ---- Write Report (member-focused)
with tabs[0]:
    st.header("Write Report")
    master = data["Incidents"]
    mode = st.radio("Mode", ["New","Edit"], horizontal=True, key="mode_write")
    defaults = {}; selected = None

    # Only show user's own drafts/submissions when editing (unless Admin)
    if mode == "Edit":
        if user["role"] in ["Admin","Reviewer"]:
            options_df = master
        else:
            options_df = master[master["CreatedBy"].astype(str) == user["username"]]
        options = options_df[PRIMARY_KEY].dropna().astype(str).tolist() if PRIMARY_KEY in options_df.columns else []
        selected = st.selectbox("Select IncidentNumber", options=options, index=None, placeholder="Choose...", key="pick_edit_write")
        if selected:
            defaults = master[master[PRIMARY_KEY].astype(str) == selected].iloc[0].to_dict()

    # Details
    with st.container(border=True):
        st.subheader("Incident Details")
        c1, c2, c3 = st.columns(3)
        inc_num = c1.text_input("IncidentNumber", value=str(defaults.get(PRIMARY_KEY,"")) if defaults else "", key="w_inc_num")
        inc_date = c2.date_input("IncidentDate", value=pd.to_datetime(defaults.get("IncidentDate")).date() if defaults.get("IncidentDate") is not None and str(defaults.get("IncidentDate")) != "NaT" else date.today(), key="w_inc_date")
        inc_time = c3.text_input("IncidentTime (HH:MM)", value=str(defaults.get("IncidentTime","")) if defaults else "", key="w_inc_time")
        c4, c5, c6 = st.columns(3)
        inc_type = c4.selectbox("IncidentType", options=[""]+lookups.get("IncidentType", []), index=0, key="w_type")
        inc_prio = c5.selectbox("ResponsePriority", options=[""]+lookups.get("ResponsePriority", []), index=0, key="w_prio")
        inc_alarm = c6.selectbox("AlarmLevel", options=[""]+lookups.get("AlarmLevel", []), index=0, key="w_alarm")
        c7, c8, c9 = st.columns(3)
        loc_name = c7.text_input("LocationName", value=str(defaults.get("LocationName","")) if defaults else "", key="w_locname")
        addr = c8.text_input("Address", value=str(defaults.get("Address","")) if defaults else "", key="w_addr")
        city = c9.text_input("City", value=str(defaults.get("City","")) if defaults else "", key="w_city")
        c10, c11, c12 = st.columns(3)
        state = c10.selectbox("State", options=[""]+lookups.get("State", []), index=0, key="w_state")
        postal = c11.text_input("PostalCode", value=str(defaults.get("PostalCode","")) if defaults else "", key="w_postal")
        shift = c12.text_input("Shift", value=str(defaults.get("Shift","")) if defaults else "", key="w_shift")

    # Narrative (big)
    with st.container(border=True):
        st.subheader("Narrative")
        narrative = st.text_area("Write full narrative here", value=str(defaults.get("Narrative","")) if defaults else "", height=300, key="w_narrative")

    # Members on Scene (multi-add box)
    with st.container(border=True):
        st.subheader("All Members on Scene")
        roster_people = data["Personnel"]
        people_opts = roster_people["Name"].dropna().astype(str).tolist() if "Name" in roster_people.columns else []
        picked_people = st.multiselect("Pick members", options=people_opts, key="w_pick_people")
        roles = lookups.get("Role", ["OIC","Driver","Firefighter"])
        c = st.columns(3)
        role_default = c[0].selectbox("Default Role", options=roles, index=0 if roles else None, key="w_role_default")
        hours_default = c[1].number_input("Default Hours", value=0.0, min_value=0.0, step=0.5, key="w_hours_default")
        if c[2].button("Add Selected Members", key="w_add_people_btn"):
            if inc_num:
                bulk_add_personnel(data, inc_num, picked_people, role_default, hours_default)
                autosave_if_enabled(data, file_path)
                st.success(f"Added {len(picked_people)} member(s).")
        # current list
        cur_per = ensure_columns(data.get("Incident_Personnel", pd.DataFrame()), CHILD_TABLES["Incident_Personnel"])
        cur_view = cur_per[cur_per[PRIMARY_KEY].astype(str) == str(inc_num)]
        st.write(f"**Total Personnel on Scene:** {len(cur_view) if not cur_view.empty else 0}")
        st.dataframe(cur_view, use_container_width=True, hide_index=True)

    # Apparatus on Scene (multi-add box)
    with st.container(border=True):
        st.subheader("Apparatus on Scene")
        roster_units = data["Apparatus"]
        unit_label = "UnitNumber" if "UnitNumber" in roster_units.columns else ("CallSign" if "CallSign" in roster_units.columns else None)
        unit_opts = roster_units[unit_label].dropna().astype(str).tolist() if unit_label else []
        picked_units = st.multiselect("Pick apparatus units", options=unit_opts, key="w_pick_units")
        c = st.columns(3)
        unit_type = c[0].selectbox("UnitType", options=[""]+lookups.get("UnitType", []), index=0, key="w_unit_type")
        unit_role = c[1].selectbox("Role", options=["Primary","Support","Water Supply","Staging"], index=0, key="w_unit_role")
        if c[2].button("Add Selected Units", key="w_add_units_btn"):
            if inc_num:
                bulk_add_units(data, inc_num, picked_units, unit_role, unit_type, [])
                autosave_if_enabled(data, file_path)
                st.success(f"Added {len(picked_units)} unit(s).")
        # current list
        cur_app = ensure_columns(data.get("Incident_Apparatus", pd.DataFrame()), CHILD_TABLES["Incident_Apparatus"])
        cur_app_view = cur_app[cur_app[PRIMARY_KEY].astype(str) == str(inc_num)]
        st.write(f"**Total Apparatus on Scene:** {len(cur_app_view) if not cur_app_view.empty else 0}")
        st.dataframe(cur_app_view, use_container_width=True, hide_index=True)

    # Times (quick)
    with st.container(border=True):
        st.subheader("Times (optional)")
        c = st.columns(4)
        alarm = c[0].text_input("Alarm (HH:MM)", key="w_alarm_time")
        enroute = c[1].text_input("Enroute (HH:MM)", key="w_enroute_time")
        arrival = c[2].text_input("Arrival (HH:MM)", key="w_arrival_time")
        clear = c[3].text_input("Clear (HH:MM)", key="w_clear_time")
        if st.button("Save Times", key="w_save_times"):
            times = data["Incident_Times"]
            new = {PRIMARY_KEY: inc_num, "Alarm": alarm, "Enroute": enroute, "Arrival": arrival, "Clear": clear}
            data["Incident_Times"] = upsert_row(times, new, key=PRIMARY_KEY)
            autosave_if_enabled(data, file_path)
            st.success("Times saved.")

    # Save/Submit buttons
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
        "CreatedBy": user["username"],
    }
    actions = st.columns(3)
    if actions[0].button("Save Draft", key="w_save_draft"):
        row_vals["Status"] = "Draft"
        data["Incidents"] = upsert_row(data["Incidents"], row_vals, key=PRIMARY_KEY)
        autosave_if_enabled(data, file_path)
        st.success("Draft saved.")
    if actions[1].button("Submit for Review", key="w_submit_review"):
        row_vals["Status"] = "Submitted"
        row_vals["SubmittedAt"] = datetime.now().strftime("%Y-%m-%d %H:%M")
        data["Incidents"] = upsert_row(data["Incidents"], row_vals, key=PRIMARY_KEY)
        autosave_if_enabled(data, file_path)
        st.success("Submitted for review.")
    if selected and actions[2].button("Open Printable", key="w_open_print"):
        with st.expander("Printable Report", expanded=True):
            printable_incident(data, selected)

# ---- Review Queue
with tabs[1]:
    st.header("Review Queue")
    pending = data["Incidents"][data["Incidents"]["Status"].astype(str) == "Submitted"]
    st.dataframe(pending, use_container_width=True, hide_index=True)
    sel = None
    if not pending.empty:
        sel = st.selectbox("Pick an Incident to review", options=pending[PRIMARY_KEY].astype(str).tolist(), index=None, placeholder="Choose...", key="pick_review_queue")
    if sel:
        with st.expander("Printable View", expanded=True):
            printable_incident(data, sel)
        comments = st.text_area("Reviewer Comments", key="rev_comments_queue")
        c = st.columns(3)
        if c[0].button("Approve", key="btn_approve_queue"):
            row = data["Incidents"][data["Incidents"][PRIMARY_KEY].astype(str) == sel].iloc[0].to_dict()
            row["Status"] = "Approved"; row["ReviewedBy"] = user["username"]; row["ReviewedAt"] = datetime.now().strftime("%Y-%m-%d %H:%M"); row["ReviewerComments"] = comments
            data["Incidents"] = upsert_row(data["Incidents"], row, key=PRIMARY_KEY); autosave_if_enabled(data, file_path); st.success("Approved.")
        if c[1].button("Reject", key="btn_reject_queue"):
            row = data["Incidents"][data["Incidents"][PRIMARY_KEY].astype(str) == sel].iloc[0].to_dict()
            row["Status"] = "Rejected"; row["ReviewedBy"] = user["username"]; row["ReviewedAt"] = datetime.now().strftime("%Y-%m-%d %H:%M"); row["ReviewerComments"] = comments or "Please revise."
            data["Incidents"] = upsert_row(data["Incidents"], row, key=PRIMARY_KEY); autosave_if_enabled(data, file_path); st.warning("Rejected.")
        if c[2].button("Send back to Draft", key="btn_backtodraft_queue"):
            row = data["Incidents"][data["Incidents"][PRIMARY_KEY].astype(str) == sel].iloc[0].to_dict()
            row["Status"] = "Draft"; row["ReviewerComments"] = comments
            data["Incidents"] = upsert_row(data["Incidents"], row, key=PRIMARY_KEY); autosave_if_enabled(data, file_path); st.info("Moved back to Draft.")

# ---- Browse
with tabs[2]:
    st.header("Browse & Filter")
    base = data["Incidents"].copy()
    fc1, fc2, fc3, fc4 = st.columns(4)
    with fc1:
        val_type = st.selectbox("IncidentType", options=[""]+lookups.get("IncidentType", []), key="filter_type")
        if val_type and "IncidentType" in base.columns:
            base = base[base["IncidentType"].astype(str) == val_type]
    with fc2:
        val_prio = st.selectbox("ResponsePriority", options=[""]+lookups.get("ResponsePriority", []), key="filter_prio")
        if val_prio and "ResponsePriority" in base.columns:
            base = base[base["ResponsePriority"].astype(str) == val_prio]
    with fc3:
        city_f = st.text_input("City contains", key="filter_city")
        if city_f and "City" in base.columns:
            base = base[base["City"].astype(str).str.contains(city_f, case=False, na=False)]
    with fc4:
        dr = st.date_input("Date range", [], key="filter_dates")
        if isinstance(dr, list) and len(dr) == 2 and "IncidentDate" in base.columns:
            start, end = pd.to_datetime(dr[0]), pd.to_datetime(dr[1])
            base = base[(pd.to_datetime(base["IncidentDate"], errors="coerce") >= start) & (pd.to_datetime(base["IncidentDate"], errors="coerce") <= end)]
    st.dataframe(base, use_container_width=True, hide_index=True)

# ---- Rosters
with tabs[3]:
    st.header("Rosters (Master Lists)")
    personnel = ensure_columns(data.get("Personnel", pd.DataFrame()), PERSONNEL_SCHEMA)
    personnel_edit = st.data_editor(personnel, num_rows="dynamic", use_container_width=True, key="editor_personnel_roster")
    if st.button("Save Personnel Roster", key="save_personnel_roster"):
        data["Personnel"] = personnel_edit; autosave_if_enabled(data, file_path); st.success("Personnel roster saved.")
    apparatus = ensure_columns(data.get("Apparatus", pd.DataFrame()), APPARATUS_SCHEMA)
    apparatus_edit = st.data_editor(apparatus, num_rows="dynamic", use_container_width=True, key="editor_apparatus_roster")
    if st.button("Save Apparatus Roster", key="save_apparatus_roster"):
        data["Apparatus"] = apparatus_edit; autosave_if_enabled(data, file_path); st.success("Apparatus roster saved.")

# ---- Reports
with tabs[4]: analytics_reports(data)

# ---- Print
with tabs[5]:
    st.header("Print Center")
    status = st.selectbox("Filter by Status", options=["","Approved","Submitted","Draft","Rejected"], key="print_status_center")
    base = data["Incidents"].copy()
    if status: base = base[base["Status"].astype(str) == status]
    st.dataframe(base, use_container_width=True, hide_index=True)
    sel = None
    if not base.empty:
        sel = st.selectbox("Pick an Incident", options=base[PRIMARY_KEY].astype(str).tolist(), index=None, key="print_pick_center")
    if sel:
        with st.expander("Printable Report", expanded=True): printable_incident(data, sel)
        st.info("Tip: Use your browser print dialog for a clean paper copy.")

# ---- Export
with tabs[6]:
    st.header("Export")
    if st.button("Build Excel for Download", key="btn_build_export"):
        payload = save_workbook_to_bytes(data)
        st.download_button("Download Excel", data=payload, file_name="fire_incident_db_export.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key="download_export")
    if st.button("Overwrite Source File Now", key="btn_overwrite_source"):
        ok, err = save_to_path(data, file_path)
        if ok: st.success(f"Overwrote: {file_path}")
        else: st.error(f"Failed to write: {err}")
