
import os
import io
from datetime import datetime, date
from typing import Dict, List

import pandas as pd
import streamlit as st

st.set_page_config(page_title="Fire Incident DB", page_icon="🚒", layout="wide")

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

def ensure_columns(df: pd.DataFrame, cols: List[str]) -> pd.DataFrame:
    if df is None:
        df = pd.DataFrame()
    for c in cols:
        if c not in df.columns:
            df[c] = pd.NA
    return df

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
            if k not in df.columns:
                df[k] = pd.NA
            df.loc[idx, k] = v
    else:
        for k in row.keys():
            if k not in df.columns:
                df[k] = pd.NA
        df = pd.concat([df, pd.DataFrame([row])], ignore_index=True)
    return df

def render_incident_form(df_master: pd.DataFrame, lookups: Dict[str, List[str]], defaults: dict) -> dict:
    vals = {}
    prefer_first = [PRIMARY_KEY, "IncidentDate", "IncidentTime", "IncidentType", "ResponsePriority", "AlarmLevel",
                    "Shift","LocationName","Address","City","State","PostalCode","Latitude","Longitude",
                    "Status","CreatedBy","SubmittedAt","ReviewedBy","ReviewedAt","ReviewerComments"]
    columns = [c for c in prefer_first if c in df_master.columns] + [c for c in df_master.columns if c not in prefer_first]
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
            elif col in ["Status","CreatedBy","SubmittedAt","ReviewedBy","ReviewedAt","ReviewerComments"]:
                st.caption(col); st.code(str(current) if pd.notna(current) else "")
                vals[col] = current
            else:
                if col in df_master.select_dtypes(include="number").columns:
                    try: base = float(current) if pd.notna(current) else 0.0
                    except Exception: base = 0.0
                    vals[col] = st.number_input(col, value=base, key=key)
                else:
                    vals[col] = st.text_input(col, value=str(current) if pd.notna(current) else "", key=key)
    return vals

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
    inc = data.get("Incidents", pd.DataFrame())
    row = inc[inc[PRIMARY_KEY].astype(str) == str(incident_number)]
    if row.empty:
        st.warning("Incident not found."); return
    rec = row.iloc[0].to_dict()
    st.markdown("## Incident Report")
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
    if rec.get("ReviewerComments"):
        st.write(f"**ReviewerComments:** {rec.get('ReviewerComments')}")
    incident_snapshot(data, str(rec.get('IncidentNumber','')))
    for t in ["Incident_Times","Incident_Personnel","Incident_Apparatus","Incident_Actions"]:
        df = data.get(t, pd.DataFrame())
        if not df.empty and PRIMARY_KEY in df.columns:
            view = df[df[PRIMARY_KEY].astype(str) == str(rec.get('IncidentNumber',''))]
            if not view.empty:
                st.subheader(t.replace("_"," ")); st.dataframe(view, use_container_width=True, hide_index=True)

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

st.sidebar.title("🚒 Fire Incident DB")
file_path = st.sidebar.text_input("Excel path", value=DEFAULT_FILE, key="sidebar_path")
uploaded = st.sidebar.file_uploader("Upload/replace workbook (.xlsx)", type=["xlsx"], key="sidebar_upload")
if uploaded:
    with open(file_path, "wb") as f: f.write(uploaded.read())
    st.sidebar.success(f"Saved to {file_path}")

data: Dict[str, pd.DataFrame] = {}
if os.path.exists(file_path): data = load_workbook(file_path)
if not data:
    st.info("Upload or point to your Excel workbook to begin."); st.stop()

data["Incidents"] = ensure_columns(data.get("Incidents", pd.DataFrame()), [PRIMARY_KEY, "IncidentDate","IncidentTime","IncidentType","ResponsePriority","AlarmLevel","Shift","LocationName","Address","City","State","PostalCode","Status","CreatedBy","SubmittedAt","ReviewedBy","ReviewedAt","ReviewerComments"])
for t, cols in CHILD_TABLES.items():
    data[t] = ensure_columns(data.get(t, pd.DataFrame()), cols)
data["Personnel"] = ensure_columns(data.get("Personnel", pd.DataFrame()), PERSONNEL_SCHEMA)
data["Apparatus"] = ensure_columns(data.get("Apparatus", pd.DataFrame()), APPARATUS_SCHEMA)
data = ensure_default_users(data)
lookups = get_lookups(data)

users_df = data["Users"]
if "user" not in st.session_state:
    sign_in_ui(users_df); st.stop()
user = st.session_state["user"]
st.sidebar.write(f"**Logged in as:** {user['name']}  \nRole: {user['role']}")
sign_out_button()

def my_reports_ui():
    st.header("My Reports")
    mine = data["Incidents"][data["Incidents"]["CreatedBy"].astype(str) == user["username"]] if "CreatedBy" in data["Incidents"].columns else pd.DataFrame()
    st.subheader("Your Drafts & Submissions")
    st.dataframe(mine, use_container_width=True, hide_index=True)

    st.subheader("Create / Edit")
    mode = st.radio("Mode", ["New Draft","Edit Existing"], horizontal=True, key="mode_my")
    defaults = {}; selected = None
    if mode == "Edit Existing" and not mine.empty:
        options = mine[PRIMARY_KEY].dropna().astype(str).tolist()
        selected = st.selectbox("Select IncidentNumber", options=options, index=None, placeholder="Choose...", key="pick_my_edit")
        if selected:
            defaults = data["Incidents"][data["Incidents"][PRIMARY_KEY].astype(str) == selected].iloc[0].to_dict()

    vals = render_incident_form(data["Incidents"], lookups, defaults)
    if "IncidentDate" in vals: vals["IncidentDate"] = pd.to_datetime(vals["IncidentDate"])
    vals["CreatedBy"] = user["username"]
    if not vals.get("Status"): vals["Status"] = "Draft"

    c = st.columns(3)
    if c[0].button("Save Draft", key="btn_save_draft"):
        data["Incidents"] = upsert_row(data["Incidents"], vals, key=PRIMARY_KEY); st.success("Draft saved.")
    if c[1].button("Submit for Review", key="btn_submit_report"):
        vals["Status"] = "Submitted"; vals["SubmittedAt"] = datetime.now().strftime("%Y-%m-%d %H:%M")
        data["Incidents"] = upsert_row(data["Incidents"], vals, key=PRIMARY_KEY); st.success("Submitted for review.")
    if selected and c[2].button("Open Printable", key="btn_my_print"):
        with st.expander("Printable Report", expanded=True): printable_incident(data, selected)

def review_queue_ui():
    st.header("Review Queue")
    pending = data["Incidents"][data["Incidents"]["Status"].astype(str) == "Submitted"]
    st.subheader("Submitted Reports"); st.dataframe(pending, use_container_width=True, hide_index=True)
    sel = None
    if not pending.empty:
        sel = st.selectbox("Pick an Incident to review", options=pending[PRIMARY_KEY].astype(str).tolist(), index=None, placeholder="Choose...", key="pick_review")
    if sel:
        st.subheader(f"Reviewing Incident #{sel}")
        with st.expander("Printable View", expanded=False): printable_incident(data, sel)
        comments = st.text_area("Reviewer Comments", key="rev_comments")
        c = st.columns(3)
        if c[0].button("Approve", key="btn_approve"):
            row = data["Incidents"][data["Incidents"][PRIMARY_KEY].astype(str) == sel].iloc[0].to_dict()
            row["Status"] = "Approved"; row["ReviewedBy"] = user["username"]; row["ReviewedAt"] = datetime.now().strftime("%Y-%m-%d %H:%M"); row["ReviewerComments"] = comments
            data["Incidents"] = upsert_row(data["Incidents"], row, key=PRIMARY_KEY); st.success("Approved.")
        if c[1].button("Reject", key="btn_reject"):
            row = data["Incidents"][data["Incidents"][PRIMARY_KEY].astype(str) == sel].iloc[0].to_dict()
            row["Status"] = "Rejected"; row["ReviewedBy"] = user["username"]; row["ReviewedAt"] = datetime.now().strftime("%Y-%m-%d %H:%M"); row["ReviewerComments"] = comments or "Please revise."
            data["Incidents"] = upsert_row(data["Incidents"], row, key=PRIMARY_KEY); st.warning("Rejected.")
        if c[2].button("Send back to Draft", key="btn_backtodraft"):
            row = data["Incidents"][data["Incidents"][PRIMARY_KEY].astype(str) == sel].iloc[0].to_dict()
            row["Status"] = "Draft"; row["ReviewerComments"] = comments
            data["Incidents"] = upsert_row(data["Incidents"], row, key=PRIMARY_KEY); st.info("Moved back to Draft.")

def admin_ui():
    st.header("Admin — All Reports & Users")
    st.subheader("All Reports"); st.dataframe(data["Incidents"], use_container_width=True, hide_index=True)
    st.subheader("Users")
    users_edit = st.data_editor(data["Users"], num_rows="dynamic", use_container_width=True, key="edit_users")
    if st.button("Save Users", key="btn_save_users"):
        data["Users"] = ensure_columns(users_edit, USERS_SCHEMA); st.success("Users saved (in memory). Use Export to write to Excel.")

def print_center_ui():
    st.header("Print Center")
    status = st.selectbox("Filter by Status", options=["","Approved","Submitted","Draft","Rejected"], key="print_status")
    base = data["Incidents"].copy()
    if status: base = base[base["Status"].astype(str) == status]
    st.dataframe(base, use_container_width=True, hide_index=True)
    sel = None
    if not base.empty:
        sel = st.selectbox("Pick an Incident", options=base[PRIMARY_KEY].astype(str).tolist(), index=None, key="print_pick")
    if sel:
        with st.expander("Printable Report", expanded=True): printable_incident(data, sel)
        st.info("Tip: Use your browser's print dialog for a clean paper copy.")

def export_ui():
    st.header("Export")
    if st.button("Build Excel for Download", key="btn_build_export"):
        payload = save_workbook_to_bytes(data)
        st.download_button("Download Excel", data=payload, file_name="fire_incident_db_export.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key="download_export")
    if st.button("Overwrite Source File", key="btn_overwrite_source"):
        payload = save_workbook_to_bytes(data)
        with open(file_path, "wb") as f: f.write(payload)
        st.success(f"Overwrote: {file_path}")

if user["role"] == "Member":
    tabs = st.tabs(["My Reports","Print","Export"])
    with tabs[0]: my_reports_ui()
    with tabs[1]: print_center_ui()
    with tabs[2]: export_ui()
elif user["role"] == "Reviewer":
    tabs = st.tabs(["My Reports","Review Queue","Print","Export"])
    with tabs[0]: my_reports_ui()
    with tabs[1]: review_queue_ui()
    with tabs[2]: print_center_ui()
    with tabs[3]: export_ui()
else:
    tabs = st.tabs(["My Reports","Review Queue","Admin","Print","Export"])
    with tabs[0]: my_reports_ui()
    with tabs[1]: review_queue_ui()
    with tabs[2]: admin_ui()
    with tabs[3]: print_center_ui()
    with tabs[4]: export_ui()
