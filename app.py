
import os
import io
from datetime import datetime, date
from typing import Dict, List

import pandas as pd
import streamlit as st

st.set_page_config(page_title="Fire Incident DB", page_icon="🚒", layout="wide")

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
USERS_SCHEMA = ["Username","Password","Role","FullName","Active"]  # Demo only; hash in production.

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

def render_dynamic_form(df_master: pd.DataFrame, lookups: Dict[str, List[str]], defaults: dict) -> dict:
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
    times = data.get("Incident_Times", pd.DataFrame())
    if not times.empty and set(["Alarm","Arrival"]).issubset(times.columns):
        st.subheader("Response Time (Arrival - Alarm) — naive HH:MM diff")
        def to_minutes(s):
            try:
                hh, mm = map(int, str(s).split(":")); return hh*60 + mm
            except Exception: return None
        tmp = times[[PRIMARY_KEY,"Alarm","Arrival"]].copy()
        tmp["t_alarm"] = tmp["Alarm"].apply(to_minutes)
        tmp["t_arrival"] = tmp["Arrival"].apply(to_minutes)
        tmp["resp_min"] = tmp.apply(lambda r: r["t_arrival"]-r["t_alarm"] if r["t_alarm"] is not None and r["t_arrival"] is not None else None, axis=1)
        st.dataframe(tmp[[PRIMARY_KEY,"Alarm","Arrival","resp_min"]], use_container_width=True, hide_index=True)

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

# ================= Assignment UIs (Assign tab + Inline) =================
def assignment_personnel_ui(data: Dict[str, pd.DataFrame], lookups: Dict[str, List[str]], inc_id: str):
    st.subheader("Assign Personnel")
    roster = ensure_columns(data.get("Personnel", pd.DataFrame()), PERSONNEL_SCHEMA)
    name_opts = roster["Name"].dropna().astype(str).tolist() if "Name" in roster.columns else []
    with st.expander("Add personnel from roster"):
        picked = st.multiselect("Select personnel", options=name_opts, key="assign_pick_people")
        roles = lookups.get("Role", ["OIC","Driver","Firefighter"])
        role = st.selectbox("Default Role", options=roles, index=0 if roles else None, key="assign_role_people")
        hours = st.number_input("Hours (each)", value=0.0, min_value=0.0, step=0.5, key="assign_hours_people")
        if st.button("Add selected personnel", key="btn_assign_people_add"):
            df = ensure_columns(data.get("Incident_Personnel", pd.DataFrame()), CHILD_TABLES["Incident_Personnel"])
            new = [{PRIMARY_KEY: inc_id, "Name": n, "Role": role, "Hours": hours} for n in picked]
            if new:
                data["Incident_Personnel"] = pd.concat([df, pd.DataFrame(new)], ignore_index=True); st.success(f"Added {len(new)} personnel.")
    df_cur = ensure_columns(data.get("Incident_Personnel", pd.DataFrame()), CHILD_TABLES["Incident_Personnel"])
    view = df_cur[df_cur[PRIMARY_KEY].astype(str) == str(inc_id)].copy()
    if not view.empty:
        view = view.reset_index(drop=False).rename(columns={"index":"_row_index"})
        st.dataframe(view, use_container_width=True, hide_index=True)
        rm_indices = st.multiselect("Select rows to remove (_row_index)", options=view["_row_index"].tolist(), key="assign_people_remove")
        if st.button("Remove selected personnel", key="btn_assign_people_remove"):
            df_cur = df_cur.drop(index=rm_indices, errors="ignore")
            data["Incident_Personnel"] = df_cur.reset_index(drop=True); st.success("Removed selected personnel.")

def assignment_apparatus_ui(data: Dict[str, pd.DataFrame], lookups: Dict[str, List[str]], inc_id: str):
    st.subheader("Assign Apparatus")
    roster = ensure_columns(data.get("Apparatus", pd.DataFrame()), APPARATUS_SCHEMA)
    label_col = "UnitNumber" if "UnitNumber" in roster.columns else ("CallSign" if "CallSign" in roster.columns else None)
    unit_opts = roster[label_col].dropna().astype(str).tolist() if label_col else []
    with st.expander("Add apparatus from roster"):
        picked = st.multiselect("Select apparatus", options=unit_opts, key="assign_pick_units")
        roles = ["Primary","Support","Water Supply","Staging"]
        role = st.selectbox("Default Role", options=roles, index=0, key="assign_role_units")
        ut = st.selectbox("UnitType (optional)", options=lookups.get("UnitType", []), index=None, placeholder="Select...", key="assign_unittype_units")
        actions = st.multiselect("Actions (optional)", options=lookups.get("Action", []), key="assign_actions_units")
        if st.button("Add selected apparatus", key="btn_assign_units_add"):
            df = ensure_columns(data.get("Incident_Apparatus", pd.DataFrame()), CHILD_TABLES["Incident_Apparatus"])
            new = [{PRIMARY_KEY: inc_id, "Unit": u, "UnitType": ut, "Role": role, "Actions": "; ".join(actions) if actions else ""} for u in picked]
            if new:
                data["Incident_Apparatus"] = pd.concat([df, pd.DataFrame(new)], ignore_index=True); st.success(f"Added {len(new)} apparatus rows.")
    df_cur = ensure_columns(data.get("Incident_Apparatus", pd.DataFrame()), CHILD_TABLES["Incident_Apparatus"])
    view = df_cur[df_cur[PRIMARY_KEY].astype(str) == str(inc_id)].copy()
    if not view.empty:
        view = view.reset_index(drop=False).rename(columns={"index":"_row_index"})
        st.dataframe(view, use_container_width=True, hide_index=True)
        rm_indices = st.multiselect("Select rows to remove (_row_index)", options=view["_row_index"].tolist(), key="assign_units_remove")
        if st.button("Remove selected apparatus", key="btn_assign_units_remove"):
            df_cur = df_cur.drop(index=rm_indices, errors="ignore")
            data["Incident_Apparatus"] = df_cur.reset_index(drop=True); st.success("Removed selected apparatus.")

# ================= App Boot =================
st.sidebar.title("🚒 Fire Incident DB")
file_path = st.sidebar.text_input("Excel path", value=DEFAULT_FILE, key="sidebar_path")
uploaded = st.sidebar.file_uploader("Upload/replace workbook (.xlsx)", type=["xlsx"], key="sidebar_upload")
if uploaded:
    with open(file_path, "wb") as f: f.write(uploaded.read())
    st.sidebar.success(f"Saved to {file_path}")

data: Dict[str, pd.DataFrame] = {}
if os.path.exists(file_path): data = load_workbook(file_path)
if not data: st.info("Upload or point to your Excel workbook to begin."); st.stop()

# Ensure tables exist
ensure_table(data, "Incidents", [PRIMARY_KEY,"IncidentDate","IncidentTime","IncidentType","ResponsePriority","AlarmLevel","Shift","LocationName","Address","City","State","PostalCode","Latitude","Longitude","Status","CreatedBy","SubmittedAt","ReviewedBy","ReviewedAt","ReviewerComments"])
for t, cols in CHILD_TABLES.items(): ensure_table(data, t, cols)
ensure_table(data, "Personnel", PERSONNEL_SCHEMA)
ensure_table(data, "Apparatus", APPARATUS_SCHEMA)
data = ensure_default_users(data)
lookups = get_lookups(data)

# Auth gate
users_df = data["Users"]
if "user" not in st.session_state:
    sign_in_ui(users_df); st.stop()
user = st.session_state["user"]
st.sidebar.write(f"**Logged in as:** {user['name']}  \\nRole: {user['role']}")
sign_out_button()

# ================= Tabs (Everything included) =================
tabs = st.tabs([
    "Browse",
    "Add / Edit Incident",
    "Assign to Incident",
    "Rosters",
    "Reports",
    "Review Queue",
    "Admin",
    "Print Center",
    "Export",
])

# ---- Browse
with tabs[0]:
    st.header("Browse & Filter Incidents")
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
        city = st.text_input("City contains", key="filter_city")
        if city and "City" in base.columns:
            base = base[base["City"].astype(str).str.contains(city, case=False, na=False)]
    with fc4:
        dr = st.date_input("Date range", [], key="filter_dates")
        if isinstance(dr, list) and len(dr) == 2 and "IncidentDate" in base.columns:
            start, end = pd.to_datetime(dr[0]), pd.to_datetime(dr[1])
            base = base[(pd.to_datetime(base["IncidentDate"], errors="coerce") >= start) & (pd.to_datetime(base["IncidentDate"], errors="coerce") <= end)]
    st.dataframe(base, use_container_width=True, hide_index=True)

# ---- Add / Edit Incident (inline assignment subtabs)
with tabs[1]:
    st.header("Add / Edit Incident (with inline assignments)")
    master = data["Incidents"]
    mode = st.radio("Mode", ["Add","Edit"], horizontal=True, key="mode_incident")
    defaults = {}; selected = None
    if mode == "Edit" and not master.empty and PRIMARY_KEY in master.columns:
        options = master[PRIMARY_KEY].dropna().astype(str).tolist()
        selected = st.selectbox("Select IncidentNumber", options=options, index=None, placeholder="Choose...", key="pick_incident_edit")
        if selected:
            defaults = master[master[PRIMARY_KEY].astype(str) == selected].iloc[0].to_dict()
    dtab, ptab, atab, ttab, prevtab = st.tabs(["Details","Personnel","Apparatus","Times/Actions","Preview"])
    with dtab:
        vals = render_dynamic_form(master, lookups, defaults)
        if "IncidentDate" in vals: vals["IncidentDate"] = pd.to_datetime(vals["IncidentDate"])
        # Auto-stamp creator for new
        if not vals.get("CreatedBy"): vals["CreatedBy"] = user["username"]
        if not vals.get("Status"): vals["Status"] = "Draft"
        if st.button("Save Incident", key="btn_save_incident"):
            data["Incidents"] = upsert_row(master, vals, key=PRIMARY_KEY); st.success("Saved.")
            if mode == "Edit" and selected is None: selected = str(vals.get(PRIMARY_KEY, ""))
    with ptab:
        st.subheader("Assign Personnel to this Incident")
        if mode == "Edit" and not selected:
            st.info("Select an IncidentNumber in Edit mode to assign personnel.")
        else:
            inc_id = selected if mode == "Edit" else vals.get(PRIMARY_KEY)
            if not inc_id: st.warning("Enter an IncidentNumber in Details and save first.")
            else:
                roster = data["Personnel"]; roles = lookups.get("Role", ["OIC","Driver","Firefighter"])
                c = st.columns(4)
                name = c[0].selectbox("Name (from roster)", options=[""]+roster["Name"].dropna().astype(str).tolist(), key="assign_person_name_inline")
                role = c[1].selectbox("Role", options=roles, index=0 if roles else None, key="assign_person_role_inline")
                hours = c[2].number_input("Hours", value=0.0, min_value=0.0, step=0.5, key="assign_person_hours_inline")
                if c[3].button("Add", key="btn_add_one_person_inline") and name:
                    df = data["Incident_Personnel"]; new_row = {PRIMARY_KEY: inc_id, "Name": name, "Role": role, "Hours": hours}
                    data["Incident_Personnel"] = pd.concat([df, pd.DataFrame([new_row])], ignore_index=True); st.success(f"Added {name}.")
                st.divider()
                view = data["Incident_Personnel"]; sub = view[view[PRIMARY_KEY].astype(str) == str(inc_id)].reset_index(drop=False).rename(columns={"index":"_row_index"})
                st.dataframe(sub, use_container_width=True, hide_index=True)
                rm = st.multiselect("Select personnel rows to remove (_row_index)", options=sub["_row_index"].tolist(), key="rm_person_rows_inline")
                if st.button("Remove selected personnel", key="btn_rm_person_rows_inline"):
                    data["Incident_Personnel"] = view.drop(index=rm, errors="ignore").reset_index(drop=True); st.success("Removed.")
    with atab:
        st.subheader("Assign Apparatus to this Incident")
        if mode == "Edit" and not selected:
            st.info("Select an IncidentNumber in Edit mode to assign apparatus.")
        else:
            inc_id = selected if mode == "Edit" else vals.get(PRIMARY_KEY)
            if not inc_id: st.warning("Enter an IncidentNumber in Details and save first.")
            else:
                roster = data["Apparatus"]
                unit_label = "UnitNumber" if "UnitNumber" in roster.columns else ("CallSign" if "CallSign" in roster.columns else None)
                unit_opts = [""] + (roster[unit_label].dropna().astype(str).tolist() if unit_label else [])
                c = st.columns(5)
                unit = c[0].selectbox("Unit (from roster)", options=unit_opts, key="assign_unit_unit_inline")
                unit_type = c[1].selectbox("UnitType", options=[""]+lookups.get("UnitType", []), index=0, key="assign_unit_type_inline")
                role = c[2].selectbox("Role", options=["Primary","Support","Water Supply","Staging"], index=0, key="assign_unit_role_inline")
                actions = c[3].multiselect("Actions", options=lookups.get("Action", []), key="assign_unit_actions_inline")
                if c[4].button("Add", key="btn_add_one_unit_inline") and unit:
                    df = data["Incident_Apparatus"]; new_row = {PRIMARY_KEY: inc_id, "Unit": unit, "UnitType": (unit_type if unit_type else None), "Role": role, "Actions": "; ".join(actions) if actions else ""}
                    data["Incident_Apparatus"] = pd.concat([df, pd.DataFrame([new_row])], ignore_index=True); st.success(f"Added {unit}.")
                st.divider()
                view = data["Incident_Apparatus"]; sub = view[view[PRIMARY_KEY].astype(str) == str(inc_id)].reset_index(drop=False).rename(columns={"index":"_row_index"})
                st.dataframe(sub, use_container_width=True, hide_index=True)
                rm = st.multiselect("Select apparatus rows to remove (_row_index)", options=sub["_row_index"].tolist(), key="rm_unit_rows_inline")
                if st.button("Remove selected apparatus", key="btn_rm_unit_rows_inline"):
                    data["Incident_Apparatus"] = view.drop(index=rm, errors="ignore").reset_index(drop=True); st.success("Removed.")
    with ttab:
        st.subheader("Times & Actions (optional)")
        inc_id = selected if mode == "Edit" else vals.get(PRIMARY_KEY)
        if not inc_id: st.info("Select or create an incident first.")
        else:
            times = data["Incident_Times"]; c = st.columns(4)
            alarm = c[0].text_input("Alarm (HH:MM)", key="time_alarm_inline")
            enroute = c[1].text_input("Enroute (HH:MM)", key="time_enroute_inline")
            arrival = c[2].text_input("Arrival (HH:MM)", key="time_arrival_inline")
            clear = c[3].text_input("Clear (HH:MM)", key="time_clear_inline")
            if st.button("Save Times", key="btn_save_times_inline"):
                new = {PRIMARY_KEY: inc_id, "Alarm": alarm, "Enroute": enroute, "Arrival": arrival, "Clear": clear}
                data["Incident_Times"] = upsert_row(times, new, key=PRIMARY_KEY); st.success("Times saved.")
            actions_list = lookups.get("Action", []); ac = st.columns(3)
            action = ac[0].selectbox("Action", options=[""]+actions_list, key="action_pick_inline")
            notes = ac[1].text_input("Notes", key="action_notes_inline")
            if ac[2].button("Add Action", key="btn_add_action_inline"):
                df = data["Incident_Actions"]; row = {PRIMARY_KEY: inc_id, "Action": action, "Notes": notes}
                data["Incident_Actions"] = pd.concat([df, pd.DataFrame([row])], ignore_index=True); st.success("Action added.")
            st.subheader("Current Actions")
            st.dataframe(data["Incident_Actions"][data["Incident_Actions"][PRIMARY_KEY].astype(str) == str(inc_id)], use_container_width=True, hide_index=True)
    with prevtab:
        if mode == "Edit" and selected:
            with st.expander("Printable Incident Report", expanded=True): printable_incident(data, selected)
        else:
            st.info("Select an incident in Edit mode to preview report.")

# ---- Assign to Incident (bulk add/remove)
with tabs[2]:
    st.header("Assign to Incident (Bulk)")
    if data["Incidents"].empty: st.info("Add an incident first.")
    else:
        inc_id = st.selectbox("IncidentNumber", options=data["Incidents"][PRIMARY_KEY].dropna().astype(str).tolist(), index=None, key="pick_incident_assign_tab")
        if inc_id:
            c1, c2 = st.columns(2)
            with c1: assignment_personnel_ui(data, lookups, inc_id)
            with c2: assignment_apparatus_ui(data, lookups, inc_id)

# ---- Rosters
with tabs[3]:
    st.header("Rosters (Master Lists)")
    st.write("Edit your master personnel and apparatus lists here.")
    personnel = ensure_columns(data.get("Personnel", pd.DataFrame()), PERSONNEL_SCHEMA)
    personnel_edit = st.data_editor(personnel, num_rows="dynamic", use_container_width=True, key="editor_personnel_roster")
    if st.button("Save Personnel Roster", key="save_personnel_roster"):
        data["Personnel"] = personnel_edit; st.success("Personnel roster saved (in memory). Use Export to write to Excel.")
    apparatus = ensure_columns(data.get("Apparatus", pd.DataFrame()), APPARATUS_SCHEMA)
    apparatus_edit = st.data_editor(apparatus, num_rows="dynamic", use_container_width=True, key="editor_apparatus_roster")
    if st.button("Save Apparatus Roster", key="save_apparatus_roster"):
        data["Apparatus"] = apparatus_edit; st.success("Apparatus roster saved (in memory). Use Export to write to Excel.")

# ---- Reports
with tabs[4]: analytics_reports(data)

# ---- Review Queue
with tabs[5]:
    st.header("Review Queue")
    pending = data["Incidents"][data["Incidents"]["Status"].astype(str) == "Submitted"]
    st.subheader("Submitted Reports"); st.dataframe(pending, use_container_width=True, hide_index=True)
    sel = None
    if not pending.empty:
        sel = st.selectbox("Pick an Incident to review", options=pending[PRIMARY_KEY].astype(str).tolist(), index=None, placeholder="Choose...", key="pick_review_queue")
    if sel:
        with st.expander("Printable View", expanded=False): printable_incident(data, sel)
        comments = st.text_area("Reviewer Comments", key="rev_comments_queue")
        c = st.columns(3)
        if c[0].button("Approve", key="btn_approve_queue"):
            row = data["Incidents"][data["Incidents"][PRIMARY_KEY].astype(str) == sel].iloc[0].to_dict()
            row["Status"] = "Approved"; row["ReviewedBy"] = user["username"]; row["ReviewedAt"] = datetime.now().strftime("%Y-%m-%d %H:%M"); row["ReviewerComments"] = comments
            data["Incidents"] = upsert_row(data["Incidents"], row, key=PRIMARY_KEY); st.success("Approved.")
        if c[1].button("Reject", key="btn_reject_queue"):
            row = data["Incidents"][data["Incidents"][PRIMARY_KEY].astype(str) == sel].iloc[0].to_dict()
            row["Status"] = "Rejected"; row["ReviewedBy"] = user["username"]; row["ReviewedAt"] = datetime.now().strftime("%Y-%m-%d %H:%M"); row["ReviewerComments"] = comments or "Please revise."
            data["Incidents"] = upsert_row(data["Incidents"], row, key=PRIMARY_KEY); st.warning("Rejected.")
        if c[2].button("Send back to Draft", key="btn_backtodraft_queue"):
            row = data["Incidents"][data["Incidents"][PRIMARY_KEY].astype(str) == sel].iloc[0].to_dict()
            row["Status"] = "Draft"; row["ReviewerComments"] = comments
            data["Incidents"] = upsert_row(data["Incidents"], row, key=PRIMARY_KEY); st.info("Moved back to Draft.")

# ---- Admin
with tabs[6]:
    st.header("Admin — All Reports & Users")
    st.subheader("All Reports"); st.dataframe(data["Incidents"], use_container_width=True, hide_index=True)
    st.subheader("Users")
    users_edit = st.data_editor(ensure_columns(data.get("Users", pd.DataFrame()), USERS_SCHEMA), num_rows="dynamic", use_container_width=True, key="edit_users")
    if st.button("Save Users", key="btn_save_users"):
        data["Users"] = ensure_columns(users_edit, USERS_SCHEMA); st.success("Users saved (in memory). Use Export to write to Excel.")

# ---- Print Center
with tabs[7]:
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
with tabs[8]:
    st.header("Export")
    if st.button("Build Excel for Download", key="btn_build_export"):
        payload = save_workbook_to_bytes(data)
        st.download_button("Download Excel", data=payload, file_name="fire_incident_db_export.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key="download_export")
    if st.button("Overwrite Source File", key="btn_overwrite_source"):
        payload = save_workbook_to_bytes(data)
        with open(file_path, "wb") as f: f.write(payload)
        st.success(f"Overwrote: {file_path}")
