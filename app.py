
import os
import io
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

DATE_LIKE = {"IncidentDate"}
TIME_LIKE = {"IncidentTime","Alarm","Enroute","Arrival","Clear"}

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

def build_person_name_series(df: pd.DataFrame) -> pd.Series:
    # Prefer Name, fallback to FullName, else FirstName+LastName
    candidates = None
    if "Name" in df.columns and df["Name"].notna().any():
        candidates = df["Name"].astype(str)
    elif "FullName" in df.columns and df["FullName"].notna().any():
        candidates = df["FullName"].astype(str)
    elif all(c in df.columns for c in ["FirstName","LastName"]):
        candidates = (df["FirstName"].fillna("").astype(str).str.strip() + " " + df["LastName"].fillna("").astype(str).str.strip()).str.strip()
    else:
        candidates = pd.Series([], dtype=str)
    if "Active" in df.columns and st.session_state.get("filter_active_only", True):
        mask = df["Active"].astype(str).str.lower().isin(["yes","true","1","active"])
        candidates = candidates[mask]
    candidates = candidates.dropna().map(lambda s: s.strip()).replace("", pd.NA).dropna().unique().tolist()
    return pd.Series(sorted(set(candidates)))

def build_unit_label_series(df: pd.DataFrame) -> pd.Series:
    # Prefer UnitNumber, then CallSign, then Name
    src = None
    for col in ["UnitNumber","CallSign","Name"]:
        if col in df.columns and df[col].notna().any():
            src = df[col].astype(str); break
    if src is None:
        src = pd.Series([], dtype=str)
    if "Active" in df.columns and st.session_state.get("filter_active_only", True):
        mask = df["Active"].astype(str).str.lower().isin(["yes","true","1","active"])
        src = src[mask]
    src = src.dropna().map(lambda s: s.strip()).replace("", pd.NA).dropna().unique().tolist()
    return pd.Series(sorted(set(src)))

# ---------------- App Boot ----------------
st.sidebar.title("📝 Fire Incident Reports")
file_path = st.sidebar.text_input("Excel path", value=DEFAULT_FILE, key="sidebar_path")
uploaded = st.sidebar.file_uploader("Upload/replace workbook (.xlsx)", type=["xlsx"], key="sidebar_upload")
if uploaded:
    with open(file_path, "wb") as f: f.write(uploaded.read())
    st.sidebar.success(f"Saved to {file_path}")
st.session_state.setdefault("autosave", True)
st.session_state["autosave"] = st.sidebar.toggle("Autosave to Excel", value=True)
st.session_state.setdefault("filter_active_only", True)
st.session_state["filter_active_only"] = st.sidebar.toggle("Show ACTIVE roster only", value=True)
st.session_state.setdefault("roster_source", "app")
st.session_state["roster_source"] = st.sidebar.selectbox("Roster source", options=["app","excel"], index=0, help="Use the live in‑app roster (session) or read directly from the Excel file.")

data: Dict[str, pd.DataFrame] = {}
if os.path.exists(file_path): data = load_workbook(file_path)
if not data: st.info("Upload or point to your Excel workbook to begin."); st.stop()

# Ensure tables exist
for name, cols in {
    "Incidents":[
        PRIMARY_KEY,"IncidentDate","IncidentTime","IncidentType","ResponsePriority","AlarmLevel","Shift",
        "LocationName","Address","City","State","PostalCode","Latitude","Longitude",
        "Narrative","Status","CreatedBy","SubmittedAt","ReviewedBy","ReviewedAt","ReviewerComments"
    ],
    "Personnel":PERSONNEL_SCHEMA,
    "Apparatus":APPARATUS_SCHEMA,
    "Incident_Times":CHILD_TABLES["Incident_Times"],
    "Incident_Personnel":CHILD_TABLES["Incident_Personnel"],
    "Incident_Apparatus":CHILD_TABLES["Incident_Apparatus"],
    "Incident_Actions":CHILD_TABLES["Incident_Actions"],
}.items():
    if name not in data: data[name] = pd.DataFrame(columns=cols)
    else: data[name] = ensure_columns(data[name], cols)

# Seed session_state rosters if not present
st.session_state.setdefault("roster_personnel", data["Personnel"].copy())
st.session_state.setdefault("roster_apparatus", data["Apparatus"].copy())

lookups = get_lookups(data)

# ---------- Auth (simple built-in demo users) ----------
def ensure_default_users(data: Dict[str, pd.DataFrame]):
    users = data.get("Users", pd.DataFrame())
    if users.empty or "Username" not in users.columns:
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

data = ensure_default_users(data)
if "user" not in st.session_state:
    sign_in_ui(data["Users"]); st.stop()
user = st.session_state["user"]
st.sidebar.write(f"**Logged in as:** {user['name']}  \\nRole: {user['role']}")
sign_out_button()

# --------------- Tabs ---------------
tabs = st.tabs(["Write Report","Review Queue","Rosters","Print","Export"])

# Helper: choose roster frame (app vs excel)
def get_roster(table: str) -> pd.DataFrame:
    use_app = st.session_state.get("roster_source","app") == "app"
    if use_app:
        return st.session_state["roster_personnel"].copy() if table == "Personnel" else st.session_state["roster_apparatus"].copy()
    else:
        return data[table].copy()

# ---- Write Report
with tabs[0]:
    st.header("Write Report")
    master = data["Incidents"]
    mode = st.radio("Mode", ["New","Edit"], horizontal=True, key="mode_write")
    defaults = {}; selected = None
    if mode == "Edit":
        options_df = master if user["role"] in ["Admin","Reviewer"] else master[master["CreatedBy"].astype(str) == user["username"]]
        options = options_df[PRIMARY_KEY].dropna().astype(str).tolist() if PRIMARY_KEY in options_df.columns else []
        selected = st.selectbox("Select IncidentNumber", options=options, index=None, placeholder="Choose...", key="pick_edit_write")
        if selected:
            defaults = master[master[PRIMARY_KEY].astype(str) == selected].iloc[0].to_dict()

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

    with st.container(border=True):
        st.subheader("Narrative")
        narrative = st.text_area("Write full narrative here", value=str(defaults.get("Narrative","")) if defaults else "", height=300, key="w_narrative")

    # Personnel picker (from live app roster by default)
    with st.container(border=True):
        st.subheader("All Members on Scene")
        roster_people = ensure_columns(get_roster("Personnel"), PERSONNEL_SCHEMA)
        # Build name list
        if "Name" in roster_people.columns and roster_people["Name"].notna().any():
            names = roster_people["Name"].astype(str)
        elif "FullName" in roster_people.columns and roster_people["FullName"].notna().any():
            names = roster_people["FullName"].astype(str)
        elif all(c in roster_people.columns for c in ["FirstName","LastName"]):
            names = (roster_people["FirstName"].fillna("").astype(str).str.strip() + " " + roster_people["LastName"].fillna("").astype(str).str.strip()).str.strip()
        else:
            names = pd.Series([], dtype=str)
        if "Active" in roster_people.columns and st.session_state.get("filter_active_only", True):
            mask = roster_people["Active"].astype(str).str.lower().isin(["yes","true","1","active"])
            names = names[mask]
        name_opts = sorted(set(names.dropna().map(lambda s: s.strip()).replace("", pd.NA).dropna().unique().tolist()))
        if len(name_opts) == 0:
            st.warning("No personnel names found in the live roster. Go to **Rosters → Personnel**, add rows, and they appear here immediately.")
            st.caption(f"Detected columns: {list(roster_people.columns)} | Rows: {len(roster_people)}")
        picked_people = st.multiselect("Pick members", options=name_opts, key="w_pick_people")
        roles = get_lookups(data).get("Role", ["OIC","Driver","Firefighter"])
        c = st.columns(3)
        role_default = c[0].selectbox("Default Role", options=roles, index=0 if roles else None, key="w_role_default")
        hours_default = c[1].number_input("Default Hours", value=0.0, min_value=0.0, step=0.5, key="w_hours_default")
        if c[2].button("Add Selected Members", key="w_add_people_btn"):
            if inc_num:
                df = ensure_columns(data.get("Incident_Personnel", pd.DataFrame()), CHILD_TABLES["Incident_Personnel"])
                new = [{PRIMARY_KEY: inc_num, "Name": n, "Role": role_default, "Hours": hours_default} for n in picked_people]
                if new:
                    data["Incident_Personnel"] = pd.concat([df, pd.DataFrame(new)], ignore_index=True)
                    autosave_if_enabled(data, file_path)
                    st.success(f"Added {len(new)} member(s).")
        cur_per = ensure_columns(data.get("Incident_Personnel", pd.DataFrame()), CHILD_TABLES["Incident_Personnel"])
        cur_view = cur_per[cur_per[PRIMARY_KEY].astype(str) == str(inc_num)]
        st.write(f"**Total Personnel on Scene:** {len(cur_view) if not cur_view.empty else 0}")
        st.dataframe(cur_view, use_container_width=True, hide_index=True)

    # Apparatus picker (from live app roster by default)
    with st.container(border=True):
        st.subheader("Apparatus on Scene")
        roster_units = ensure_columns(get_roster("Apparatus"), APPARATUS_SCHEMA)
        label = None
        for col in ["UnitNumber","CallSign","Name"]:
            if col in roster_units.columns and roster_units[col].notna().any(): label = col; break
        unit_series = roster_units[label].astype(str) if label else pd.Series([], dtype=str)
        if "Active" in roster_units.columns and st.session_state.get("filter_active_only", True):
            mask = roster_units["Active"].astype(str).str.lower().isin(["yes","true","1","active"])
            unit_series = unit_series[mask]
        unit_opts = sorted(set(unit_series.dropna().map(lambda s: s.strip()).replace("", pd.NA).dropna().unique().tolist()))
        if len(unit_opts) == 0:
            st.warning("No apparatus in the live roster. Go to **Rosters → Apparatus**, add rows, and they appear here immediately.")
            st.caption(f"Detected columns: {list(roster_units.columns)} | Rows: {len(roster_units)}")
        picked_units = st.multiselect("Pick apparatus units", options=unit_opts, key="w_pick_units")
        c = st.columns(3)
        unit_type = c[0].selectbox("UnitType", options=[""]+get_lookups(data).get("UnitType", []), index=0, key="w_unit_type")
        unit_role = c[1].selectbox("Role", options=["Primary","Support","Water Supply","Staging"], index=0, key="w_unit_role")
        if c[2].button("Add Selected Units", key="w_add_units_btn"):
            if inc_num:
                df = ensure_columns(data.get("Incident_Apparatus", pd.DataFrame()), CHILD_TABLES["Incident_Apparatus"])
                new = [{PRIMARY_KEY: inc_num, "Unit": u, "UnitType": (unit_type if unit_type else None), "Role": unit_role, "Actions": ""} for u in picked_units]
                if new:
                    data["Incident_Apparatus"] = pd.concat([df, pd.DataFrame(new)], ignore_index=True)
                    autosave_if_enabled(data, file_path)
                    st.success(f"Added {len(new)} unit(s).")
        cur_app = ensure_columns(data.get("Incident_Apparatus", pd.DataFrame()), CHILD_TABLES["Incident_Apparatus"])
        cur_app_view = cur_app[cur_app[PRIMARY_KEY].astype(str) == str(inc_num)]
        st.write(f"**Total Apparatus on Scene:** {len(cur_app_view) if not cur_app_view.empty else 0}")
        st.dataframe(cur_app_view, use_container_width=True, hide_index=True)

    # Times + Save/Submit
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
            autosave_if_enabled(data, file_path); st.success("Times saved.")

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
        autosave_if_enabled(data, file_path); st.success("Draft saved.")
    if actions[1].button("Submit for Review", key="w_submit_review"):
        row_vals["Status"] = "Submitted"; row_vals["SubmittedAt"] = datetime.now().strftime("%Y-%m-%d %H:%M")
        data["Incidents"] = upsert_row(data["Incidents"], row_vals, key=PRIMARY_KEY)
        autosave_if_enabled(data, file_path); st.success("Submitted for review.")
    if selected and actions[2].button("Open Printable", key="w_open_print"):
        with st.expander("Printable Report", expanded=True):
            st.header(f"Incident Report — #{selected}")
            # simple printable
            st.json(row_vals)

# ---- Review Queue
with tabs[1]:
    st.header("Review Queue")
    pending = data["Incidents"][data["Incidents"]["Status"].astype(str) == "Submitted"]
    st.dataframe(pending, use_container_width=True, hide_index=True)
    sel = None
    if not pending.empty:
        sel = st.selectbox("Pick an Incident to review", options=pending[PRIMARY_KEY].astype(str).tolist(), index=None, placeholder="Choose...", key="pick_review_queue")
    if sel:
        st.subheader("Printable View")
        # reuse a quick printable
        rec = data["Incidents"][data["Incidents"][PRIMARY_KEY].astype(str) == sel].iloc[0].to_dict()
        st.json(rec)
        comments = st.text_area("Reviewer Comments", key="rev_comments_queue")
        c = st.columns(3)
        if c[0].button("Approve", key="btn_approve_queue"):
            row = rec; row["Status"] = "Approved"; row["ReviewedBy"] = user["username"]; row["ReviewedAt"] = datetime.now().strftime("%Y-%m-%d %H:%M"); row["ReviewerComments"] = comments
            data["Incidents"] = upsert_row(data["Incidents"], row, key=PRIMARY_KEY); autosave_if_enabled(data, file_path); st.success("Approved.")
        if c[1].button("Reject", key="btn_reject_queue"):
            row = rec; row["Status"] = "Rejected"; row["ReviewedBy"] = user["username"]; row["ReviewedAt"] = datetime.now().strftime("%Y-%m-%d %H:%M"); row["ReviewerComments"] = comments or "Please revise."
            data["Incidents"] = upsert_row(data["Incidents"], row, key=PRIMARY_KEY); autosave_if_enabled(data, file_path); st.warning("Rejected.")
        if c[2].button("Send back to Draft", key="btn_backtodraft_queue"):
            row = rec; row["Status"] = "Draft"; row["ReviewerComments"] = comments
            data["Incidents"] = upsert_row(data["Incidents"], row, key=PRIMARY_KEY); autosave_if_enabled(data, file_path); st.info("Moved back to Draft.")

# ---- Rosters
with tabs[2]:
    st.header("Rosters")
    st.caption("Edits update the in‑app roster immediately; click Save to write to Excel. The **Write Report** pickers use the in‑app roster when 'Roster source' = app.")
    # Personnel
    personnel = ensure_columns(st.session_state["roster_personnel"], PERSONNEL_SCHEMA)
    st.session_state["roster_personnel"] = st.data_editor(personnel, num_rows="dynamic", use_container_width=True, key="editor_personnel_roster")
    c = st.columns(2)
    if c[0].button("Save Personnel Roster to Excel", key="save_personnel_roster"):
        data["Personnel"] = ensure_columns(st.session_state["roster_personnel"], PERSONNEL_SCHEMA); autosave_if_enabled(data, file_path); st.success("Personnel roster saved to Excel.")
    # Apparatus
    apparatus = ensure_columns(st.session_state["roster_apparatus"], APPARATUS_SCHEMA)
    st.session_state["roster_apparatus"] = st.data_editor(apparatus, num_rows="dynamic", use_container_width=True, key="editor_apparatus_roster")
    if c[1].button("Save Apparatus Roster to Excel", key="save_apparatus_roster"):
        data["Apparatus"] = ensure_columns(st.session_state["roster_apparatus"], APPARATUS_SCHEMA); autosave_if_enabled(data, file_path); st.success("Apparatus roster saved to Excel.")

# ---- Print
with tabs[3]:
    st.header("Print")
    status = st.selectbox("Filter by Status", options=["","Approved","Submitted","Draft","Rejected"], key="print_status_center")
    base = data["Incidents"].copy()
    if status: base = base[base["Status"].astype(str) == status]
    st.dataframe(base, use_container_width=True, hide_index=True)
    sel = None
    if not base.empty:
        sel = st.selectbox("Pick an Incident", options=base[PRIMARY_KEY].astype(str).tolist(), index=None, key="print_pick_center")
    if sel:
        rec = data["Incidents"][data["Incidents"][PRIMARY_KEY].astype(str) == sel].iloc[0].to_dict()
        st.json(rec)

# ---- Export
with tabs[4]:
    st.header("Export")
    if st.button("Build Excel for Download", key="btn_build_export"):
        payload = save_workbook_to_bytes(data)
        st.download_button("Download Excel", data=payload, file_name="fire_incident_db_export.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key="download_export")
    if st.button("Overwrite Source File Now", key="btn_overwrite_source"):
        ok, err = save_to_path(data, file_path)
        if ok: st.success(f"Overwrote: {file_path}")
        else: st.error(f"Failed to write: {err}")
