
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
    "Incident_Personnel": ["IncidentNumber","PersonnelID","Name","Role","Hours","RespondedIn"],
    "Incident_Apparatus": ["IncidentNumber","ApparatusID","Unit","UnitType","Role","Actions"],
    "Incident_Actions": ["IncidentNumber","Action","Notes"],
}
PERSONNEL_SCHEMA = ["PersonnelID","Name","UnitNumber","Rank","Badge","Phone","Email","Address","City","State","PostalCode","Certifications","Active","FirstName","LastName","FullName"]
APPARATUS_SCHEMA = ["ApparatusID","UnitNumber","CallSign","UnitType","GPM","TankSize","SeatingCapacity","Station","Active","Name"]
USERS_SCHEMA = ["Username","Password","Role","FullName","Active",
                "CanWrite","CanEditOwn","CanEditAll","CanReview","CanApprove","CanManageUsers","CanEditRosters","CanPrint","CanDeleteArchive"]

LOOKUP_SHEETS = {
    "List_IncidentType": "IncidentType",
    "List_AlarmLevel": "AlarmLevel",
    "List_ResponsePriority": "ResponsePriority",
    "List_PersonnelRoles": "Role",
    "List_UnitTypes": "UnitType",
    "List_Actions": "Action",
    "List_States": "State",
}

ROLE_PRESETS = {
    "Admin":   {"CanWrite":True,"CanEditOwn":True,"CanEditAll":True,"CanReview":True,"CanApprove":True,"CanManageUsers":True,"CanEditRosters":True,"CanPrint":True,"CanDeleteArchive":True},
    "Reviewer":{"CanWrite":False,"CanEditOwn":False,"CanEditAll":False,"CanReview":True,"CanApprove":True,"CanManageUsers":False,"CanEditRosters":False,"CanPrint":True,"CanDeleteArchive":False},
    "Member":  {"CanWrite":True,"CanEditOwn":True,"CanEditAll":False,"CanReview":False,"CanApprove":False,"CanManageUsers":False,"CanEditRosters":False,"CanPrint":True,"CanDeleteArchive":False},
}

def load_workbook(path: str) -> Dict[str, pd.DataFrame]:
    try:
        xls = pd.ExcelFile(path)
        return {name.strip(): xls.parse(name) for name in xls.sheet_names}
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

# --- ID lookup helpers (used when adding roster selections to an incident) ---
def _lookup_personnel_id(personnel_df: pd.DataFrame, name: str):
    if personnel_df is None or personnel_df.empty:
        return pd.NA
    if "Name" in personnel_df.columns:
        m = personnel_df[personnel_df["Name"].astype(str) == str(name)]
        if not m.empty and "PersonnelID" in m.columns:
            return m.iloc[0]["PersonnelID"]
    # fallback: try FirstName/LastName and Rank combined if needed
    return pd.NA

def _lookup_apparatus_id(app_df: pd.DataFrame, unit_value: str):
    if app_df is None or app_df.empty:
        return pd.NA
    u = str(unit_value)
    # try matching by Name, CallSign, UnitNumber (in that order)
    for col in ["Name", "CallSign", "UnitNumber", "Unit"]:
        if col in app_df.columns:
            m = app_df[app_df[col].astype(str) == u]
            if not m.empty and "ApparatusID" in m.columns:
                return m.iloc[0]["ApparatusID"]
    return pd.NA

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

def _name_rank_first_last(row: pd.Series) -> str:
    fn = str(row.get("FirstName") or "").strip()
    ln = str(row.get("LastName") or "").strip()
    rk = str(row.get("Rank") or "").strip()
    parts = [p for p in [rk, fn, ln] if p]
    return " ".join(parts).strip()

def build_person_options(df: pd.DataFrame) -> list:
    if "Name" in df and df["Name"].notna().any():
        s = df["Name"].astype(str)
    elif "FullName" in df and df["FullName"].notna().any():
        s = df["FullName"].astype(str)
    elif all(c in df.columns for c in ["FirstName","LastName","Rank"]):
        s = df.apply(_name_rank_first_last, axis=1)
    elif all(c in df.columns for c in ["FirstName","LastName"]):
        s = (df["FirstName"].fillna("").astype(str).str.strip() + " " + df["LastName"].fillna("").astype(str).str.strip()).str.strip()
    else:
        s = pd.Series([], dtype=str)
    vals = s.dropna().map(lambda x: x.strip()).replace("", pd.NA).dropna().unique().tolist()
    return sorted(set(vals))

def build_unit_options(df: pd.DataFrame) -> list:
    for col in ["UnitNumber","CallSign","Name"]:
        if col in df.columns and df[col].notna().any():
            s = df[col].astype(str); break
    else:
        s = pd.Series([], dtype=str)
    vals = s.dropna().map(lambda x: x.strip()).replace("", pd.NA).dropna().unique().tolist()
    return sorted(set(vals))

def repair_rosters(data: Dict[str, pd.DataFrame]) -> Dict[str, pd.DataFrame]:
    p = ensure_columns(data.get("Personnel", pd.DataFrame()), PERSONNEL_SCHEMA).copy()
    if "Rank" in p.columns and not p.empty:
        p["Rank"] = p["Rank"].astype(str)  # free-text ranks
    if not p.empty:
        mask_name_blank = p["Name"].isna() | (p["Name"].astype(str).str.strip()=="")
        p.loc[mask_name_blank, "Name"] = p.loc[mask_name_blank].apply(_name_rank_first_last, axis=1)
        mask_full_blank = p["FullName"].isna() | (p["FullName"].astype(str).str.strip()=="")
        p.loc[mask_full_blank, "FullName"] = p.loc[mask_full_blank].apply(_name_rank_first_last, axis=1)
        if "Active" in p.columns:
            m = p["Active"].isna() | (p["Active"].astype(str).str.strip()=="")
            p.loc[m, "Active"] = "Yes"
        else:
            p["Active"] = "Yes"
    data["Personnel"] = p

    a = ensure_columns(data.get("Apparatus", pd.DataFrame()), APPARATUS_SCHEMA).copy()
    if not a.empty:
        if "Active" in a.columns:
            m = a["Active"].isna() | (a["Active"].astype(str).str.strip()=="")
            a.loc[m, "Active"] = "Yes"
        else:
            a["Active"] = "Yes"
    data["Apparatus"] = a
    return data

def _coerce_bool(x):
    s = str(x).strip().lower()
    return s in ("1","true","yes","y")

def apply_role_presets(df: pd.DataFrame) -> pd.DataFrame:
    df = ensure_columns(df, USERS_SCHEMA).copy()
    for i, row in df.iterrows():
        role = str(row.get("Role","")).strip() or "Member"
        preset = ROLE_PRESETS.get(role, ROLE_PRESETS["Member"])
        for k, v in preset.items():
            if pd.isna(row.get(k)) or str(row.get(k)).strip()=="":
                df.at[i, k] = v
    if "Active" in df.columns:
        mask = df["Active"].isna() | (df["Active"].astype(str).str.strip()=="")
        df.loc[mask, "Active"] = "Yes"
    return df

def can(user_row: dict, perm: str) -> bool:
    return _coerce_bool(user_row.get(perm, False))

st.sidebar.title("📝 Fire Incident Reports — v4.3.2")
file_path = st.sidebar.text_input("Excel path", value=DEFAULT_FILE, key="path_input_auth")
uploaded = st.sidebar.file_uploader("Upload/replace workbook (.xlsx)", type=["xlsx"], key="upload_auth")
if uploaded:
    with open(file_path, "wb") as f: f.write(uploaded.read())
    st.sidebar.success(f"Saved to {file_path}")
st.session_state.setdefault("autosave", True)
st.session_state["autosave"] = st.sidebar.toggle("Autosave to Excel", value=True, key="autosave_auth")
st.sidebar.caption(f"File exists: {'✅' if os.path.exists(file_path) else '❌'}")

data: Dict[str, pd.DataFrame] = {}
if os.path.exists(file_path):
    data = load_workbook(file_path)
else:
    st.info("Upload or point to your Excel workbook to begin.")
    st.stop()

ensure_table(data, "Incidents", [
    PRIMARY_KEY,"IncidentDate","IncidentTime","IncidentType","ResponsePriority","AlarmLevel","Shift",
    "LocationName","Address","City","State","PostalCode","Latitude","Longitude",
    "Narrative","Status","CreatedBy","SubmittedAt","ReviewedBy","ReviewedAt","ReviewerComments",
    "CallerName","CallerPhone","ArchiveStatus"
])
ensure_table(data, "Personnel", PERSONNEL_SCHEMA)
ensure_table(data, "Apparatus", APPARATUS_SCHEMA)
for t, cols in CHILD_TABLES.items(): ensure_table(data, t, cols)

users = ensure_columns(data.get("Users", pd.DataFrame()), USERS_SCHEMA)
if users.empty or "Username" not in users.columns or users["Username"].isna().all():
    users = pd.DataFrame([
        {"Username":"admin","Password":"admin","Role":"Admin","FullName":"Administrator","Active":"Yes"},
        {"Username":"review","Password":"review","Role":"Reviewer","FullName":"Reviewer","Active":"Yes"},
        {"Username":"member","Password":"member","Role":"Member","FullName":"Member User","Active":"Yes"},
    ])
users = apply_role_presets(users)
data["Users"] = users

def sign_in_ui(users_df: pd.DataFrame):
    st.header("Sign In")
    u = st.text_input("Username", key="login_user_auth")
    p = st.text_input("Password", type="password", key="login_pass_auth")
    ok = st.button("Sign In", key="btn_login_auth")
    if ok:
        row = users_df[(users_df["Username"].astype(str)==u) & (users_df["Password"].astype(str)==p) & (users_df["Active"].astype(str).str.lower().isin(["yes","true","1"]))]
        if not row.empty:
            st.session_state["user"] = row.iloc[0].to_dict()
            st.success(f"Welcome, {row.iloc[0].get('FullName', u)}!"); st.experimental_rerun()
        else:
            st.error("Invalid credentials or inactive user.")

def sign_out_button():
    if st.button("Sign Out", key="btn_logout_auth"):
        st.session_state.pop("user", None); st.experimental_rerun()

data = repair_rosters(data)
lookups = get_lookups(data)

if "user" not in st.session_state:
    sign_in_ui(data["Users"]); st.stop()
user = st.session_state["user"]
st.sidebar.write(f"**Logged in as:** {user.get('FullName', user.get('Username',''))}  \\nRole: {user.get('Role','')}")
sign_out_button()

tabs = st.tabs(["Write Report","Review Queue","Rejected","Approved","Archive","Rosters","Print","Export","Admin","Diagnostics"])

with tabs[0]:
    st.header("Write Report")
    # --- Clear / New Report ---
    if st.button("🆕 New Report (Clear)", key="btn_clear_write"):
        keys_to_clear = [
            k for k in list(st.session_state.keys())
            if k.startswith("w_")
            or k.startswith("editor_incident_")
            or k.startswith("pick_")
        ]
        for k in keys_to_clear:
            st.session_state.pop(k, None)

        st.session_state.pop("edit_incident_preselect", None)
        st.session_state.pop("force_edit_mode", None)
        st.experimental_rerun()

    master = data["Incidents"].copy()
    preselect = st.session_state.get("edit_incident_preselect")
    force_edit = st.session_state.get("force_edit_mode", False)
    if preselect:
        st.info(f"Editing incident sent back to Draft: {preselect}")
    mode_index = 1 if (preselect or force_edit) else 0
    mode = st.radio("Mode", ["New","Edit"], horizontal=True, index=mode_index, key="mode_write_auth")

    defaults = {}; selected = None
    if mode == "Edit" and not master.empty:
        if can(user,"CanEditAll"):
            options_df = master
        elif can(user,"CanEditOwn"):
            options_df = master[master["CreatedBy"].astype(str) == user.get("Username")]
        else:
            options_df = master.iloc(0,0)
        options = options_df[PRIMARY_KEY].dropna().astype(str).tolist() if PRIMARY_KEY in options_df.columns else []
        kwargs = {"options": options, "placeholder": "Choose...", "key": "pick_edit_write_auth"}
        if preselect and preselect in options:
            kwargs["index"] = options.index(preselect)
        selected = st.selectbox("Select IncidentNumber", **kwargs)
        if selected:
            defaults = master[master[PRIMARY_KEY].astype(str) == selected].iloc[0].to_dict()
            st.session_state["edit_incident_preselect"] = None
            st.session_state["force_edit_mode"] = False

    with st.container(border=True):
        st.subheader("Incident Details")
        c1, c2, c3 = st.columns(3)
        inc_num = c1.text_input("IncidentNumber", value=str(defaults.get(PRIMARY_KEY,"")) if defaults else "", key="w_inc_num_auth")
        inc_date = c2.date_input("IncidentDate", value=pd.to_datetime(defaults.get("IncidentDate")).date() if defaults.get("IncidentDate") is not None and str(defaults.get("IncidentDate")) != "NaT" else date.today(), key="w_inc_date_auth")
        inc_time = c3.text_input("IncidentTime (HH:MM)", value=str(defaults.get("IncidentTime","")) if defaults else "", key="w_inc_time_auth")
        c4, c5, c6 = st.columns(3)
        inc_type = c4.selectbox("IncidentType", options=[""]+lookups.get("IncidentType", []), index=([""]+lookups.get("IncidentType", [])).index(str(defaults.get("IncidentType",""))) if defaults.get("IncidentType") in lookups.get("IncidentType", []) else 0, key="w_type_auth")
        inc_prio = c5.selectbox("ResponsePriority", options=[""]+lookups.get("ResponsePriority", []), index=([""]+lookups.get("ResponsePriority", [])).index(str(defaults.get("ResponsePriority",""))) if defaults.get("ResponsePriority") in lookups.get("ResponsePriority", []) else 0, key="w_prio_auth")
        inc_alarm = c6.selectbox("AlarmLevel", options=[""]+lookups.get("AlarmLevel", []), index=([""]+lookups.get("AlarmLevel", [])).index(str(defaults.get("AlarmLevel",""))) if defaults.get("AlarmLevel") in lookups.get("AlarmLevel", []) else 0, key="w_alarm_auth")

    # Caller Information (compact)
    col_call1, col_call2 = st.columns(2)
    with col_call1:
        caller_name = st.text_input("Caller", key="caller_name")
    with col_call2:
        caller_phone = st.text_input("Caller Phone", key="caller_phone")

        c7, c8, c9 = st.columns(3)
        loc_name = c7.text_input("LocationName", value=str(defaults.get("LocationName","")) if defaults else "", key="w_locname_auth")
        addr = c8.text_input("Address", value=str(defaults.get("Address","")) if defaults else "", key="w_addr_auth")
        city = c9.text_input("City", value=str(defaults.get("City","")) if defaults else "", key="w_city_auth")
        c10, c11, c12 = st.columns(3)
        state = c10.text_input("State", value=str(defaults.get("State","")) if defaults else "", key="w_state_auth")
        postal = c11.text_input("PostalCode", value=str(defaults.get("PostalCode","")) if defaults else "", key="w_postal_auth")
        shift = c12.text_input("Shift", value=str(defaults.get("Shift","")) if defaults else "", key="w_shift_auth")

    with st.container(border=True):
        st.subheader("Narrative")
        narrative = st.text_area("Write full narrative here", value=str(defaults.get("Narrative","")) if defaults else "", height=320, key="w_narrative_auth")

    with st.container(border=True):
        st.subheader("All Members on Scene")
        people_df = ensure_columns(data.get("Personnel", pd.DataFrame()), PERSONNEL_SCHEMA)
        if "Rank" in people_df.columns:
            people_df["Rank"] = people_df["Rank"].astype(str)
        person_opts = build_person_options(people_df)
        app_df_all = ensure_columns(data.get("Apparatus", pd.DataFrame()), APPARATUS_SCHEMA)
        unit_opts_all = build_unit_options(app_df_all)
        picked_people = st.multiselect("Pick members", options=person_opts, key="w_pick_people_auth")
        roles = lookups.get("Role", ["OIC","Driver","Firefighter"])
        cc = st.columns(4)
        role_default = cc[0].selectbox("Default Role", options=roles, index=0 if roles else None, key="w_role_default_auth")
        hours_default = cc[1].number_input("Default Hours", value=0.0, min_value=0.0, step=0.5, key="w_hours_default_auth")
        responded_in_default = cc[2].selectbox("Responded In (optional)", options=[""]+unit_opts_all, index=0, key="w_resp_in_default_auth")
        if cc[3].button("Add Selected Members", key="w_add_people_btn_auth"):
            if not inc_num or str(inc_num).strip() == "":
                st.error("Enter **IncidentNumber** before adding members.")
            else:
                inc_key = str(inc_num).strip()
                df = ensure_columns(data.get("Incident_Personnel", pd.DataFrame()), CHILD_TABLES["Incident_Personnel"])
                new = []
                people_df = data.get('Personnel', pd.DataFrame())
                for n in picked_people:
                    pid = _lookup_personnel_id(people_df, n)
                    new.append({
                        PRIMARY_KEY: inc_key,
                        'PersonnelID': pid,
                        'Name': n,
                        'Role': role_default,
                        'Hours': hours_default,
                        'RespondedIn': (responded_in_default or None),
                    })
                if new:
                    data["Incident_Personnel"] = pd.concat([df, pd.DataFrame(new)], ignore_index=True)
                    if st.session_state.get("autosave", True): save_to_path(data, file_path)
                    st.success(f"Added {len(new)} member(s) to incident {inc_key}.")
                else:
                    st.warning("No members selected.")
        cur_per = ensure_columns(data.get("Incident_Personnel", pd.DataFrame()), CHILD_TABLES["Incident_Personnel"])
        this_per = cur_per[cur_per[PRIMARY_KEY].astype(str) == (str(inc_num).strip() if inc_num else "__none__")].copy()
        if not this_per.empty and "Delete" not in this_per.columns:
            this_per["Delete"] = False
        st.write(f"**Total Personnel on Scene:** {0 if this_per.empty else len(this_per)}")
        this_per_edit = st.data_editor(this_per, num_rows="dynamic", use_container_width=True, key="editor_incident_personnel")
        cdel = st.columns(2)
        if cdel[0].button("Save Personnel Grid", key="btn_save_incident_personnel"):
            base = cur_per[cur_per[PRIMARY_KEY].astype(str) != (str(inc_num).strip() if inc_num else "__none__")]
            if "Delete" in this_per_edit.columns:
                this_per_edit = this_per_edit[this_per_edit["Delete"] != True].drop(columns=["Delete"], errors="ignore")
            data["Incident_Personnel"] = pd.concat([base, this_per_edit], ignore_index=True)
            if st.session_state.get("autosave", True): save_to_path(data, file_path)
            st.success("Incident personnel updated (removals applied if any).")

    with st.container(border=True):
        st.subheader("Apparatus on Scene")
        app_df = ensure_columns(data.get("Apparatus", pd.DataFrame()), APPARATUS_SCHEMA)
        unit_opts = build_unit_options(app_df)
        picked_units = st.multiselect("Pick apparatus units", options=unit_opts, key="w_pick_units_auth")
        unit_type_options = list(dict.fromkeys(["Mini Pumper"] + lookups.get("UnitType", [])))
        cc2 = st.columns(4)
        unit_type = cc2[0].selectbox("UnitType", options=[""]+unit_type_options, index=0, key="w_unit_type_auth")
        unit_role = cc2[1].selectbox("Role", options=["Primary","Support","Water Supply","Staging"], index=0, key="w_unit_role_auth")
        unit_actions = cc2[2].text_input("Actions (e.g., 'Directing traffic')", key="w_unit_actions_auth")
        if cc2[3].button("Add Selected Units", key="w_add_units_btn_auth"):
            if not inc_num or str(inc_num).strip() == "":
                st.error("Enter **IncidentNumber** before adding apparatus.")
            else:
                inc_key = str(inc_num).strip()
                df = ensure_columns(data.get("Incident_Apparatus", pd.DataFrame()), CHILD_TABLES["Incident_Apparatus"])
                new = []
                app_df = data.get('Apparatus', pd.DataFrame())
                for u in picked_units:
                    aid = _lookup_apparatus_id(app_df, u)
                    new.append({
                        PRIMARY_KEY: inc_key,
                        'ApparatusID': aid,
                        'Unit': u,
                        'UnitType': (unit_type if unit_type else None),
                        'Role': unit_role,
                        'Actions': unit_actions or '',
                    })
                if new:
                    data["Incident_Apparatus"] = pd.concat([df, pd.DataFrame(new)], ignore_index=True)
                    if st.session_state.get("autosave", True): save_to_path(data, file_path)
                    st.success(f"Added {len(new)} unit(s) to incident {inc_key}.")
                else:
                    st.warning("No units selected.")
        cur_app = ensure_columns(data.get("Incident_Apparatus", pd.DataFrame()), CHILD_TABLES["Incident_Apparatus"])
        this_app = cur_app[cur_app[PRIMARY_KEY].astype(str) == (str(inc_num).strip() if inc_num else "__none__")].copy()
        if not this_app.empty and "Delete" not in this_app.columns:
            this_app["Delete"] = False
        st.write(f"**Total Apparatus on Scene:** {0 if this_app.empty else len(this_app)}")
        this_app_edit = st.data_editor(this_app, num_rows="dynamic", use_container_width=True, key="editor_incident_apparatus")
        cdel2 = st.columns(2)
        if cdel2[0].button("Save Apparatus Grid", key="btn_save_incident_apparatus"):
            base = cur_app[cur_app[PRIMARY_KEY].astype(str) != (str(inc_num).strip() if inc_num else "__none__")]
            if "Delete" in this_app_edit.columns:
                this_app_edit = this_app_edit[this_app_edit["Delete"] != True].drop(columns=["Delete"], errors="ignore")
            data["Incident_Apparatus"] = pd.concat([base, this_app_edit], ignore_index=True)
            if st.session_state.get("autosave", True): save_to_path(data, file_path)
            st.success("Incident apparatus updated (removals applied if any).")

    with st.container(border=True):
        st.subheader("Times (optional)")
        t1, t2, t3, t4 = st.columns(4)
        alarm = t1.text_input("Alarm (HH:MM)", key="w_alarm_time_auth")
        enroute = t2.text_input("Enroute (HH:MM)", key="w_enroute_time_auth")
        arrival = t3.text_input("Arrival (HH:MM)", key="w_arrival_time_auth")
        clear = t4.text_input("Clear (HH:MM)", key="w_clear_time_auth")
        if st.button("Save Times", key="w_save_times_auth"):
            if not inc_num or str(inc_num).strip() == "":
                st.error("Enter **IncidentNumber** before saving times.")
            else:
                inc_key = str(inc_num).strip()
                times = data["Incident_Times"]
                new = {PRIMARY_KEY: inc_key, "Alarm": alarm, "Enroute": enroute, "Arrival": arrival, "Clear": clear}
                data["Incident_Times"] = upsert_row(times, new, key=PRIMARY_KEY)
                if st.session_state.get("autosave", True): save_to_path(data, file_path)
                st.success("Times saved.")

    row_vals = {
        PRIMARY_KEY: (str(inc_num).strip() if inc_num else ""),
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
        "CallerName": caller_name,
        "CallerPhone": caller_phone,
        "Narrative": narrative,
        "CreatedBy": user.get("Username",""),
    }
    a = st.columns(3)
    if a[0].button("Save Draft", key="w_save_draft_btn"):
        if not can(user,"CanWrite"):
            st.error("You do not have permission to write.")
        elif not row_vals[PRIMARY_KEY]:
            st.error("Enter **IncidentNumber** before saving.")
        else:
            row_vals["Status"] = "Draft"
            data["Incidents"] = upsert_row(data["Incidents"], row_vals, key=PRIMARY_KEY)
            if st.session_state.get("autosave", True): save_to_path(data, file_path)
            st.success("Draft saved.")
    if a[1].button("Submit for Review", key="w_submit_review_btn"):
        if not can(user,"CanWrite"):
            st.error("You do not have permission to submit.")
        elif not row_vals[PRIMARY_KEY]:
            st.error("Enter **IncidentNumber** before submitting.")
        else:
            row_vals["Status"] = "Submitted"; row_vals["SubmittedAt"] = datetime.now().strftime("%Y-%m-%d %H:%M")
            data["Incidents"] = upsert_row(data["Incidents"], row_vals, key=PRIMARY_KEY)
            if st.session_state.get("autosave", True): save_to_path(data, file_path)
            st.success("Submitted for review.")

with tabs[1]:
    st.header("Review Queue")
    pending = data["Incidents"][data["Incidents"]["Status"].astype(str) == "Submitted"]
    st.dataframe(pending, use_container_width=True, hide_index=True, key="grid_pending_auth")
    sel = None
    if not pending.empty:
        sel = st.selectbox("Pick an Incident to review", options=pending[PRIMARY_KEY].astype(str).tolist(), index=None, placeholder="Choose...", key="pick_review_queue_auth")
    if sel:
        rec = data["Incidents"][data["Incidents"][PRIMARY_KEY].astype(str) == sel].iloc[0].to_dict()
        st.subheader(f"Incident {sel}")
        st.write(f"**Type:** {rec.get('IncidentType','')}  |  **Priority:** {rec.get('ResponsePriority','')}  |  **Alarm:** {rec.get('AlarmLevel','')}")
        st.write(f"**Location:** {rec.get('Address','')} {rec.get('City','')} {rec.get('State','')} {rec.get('PostalCode','')}")
        st.write("**Narrative:**")
        st.text_area("Narrative (read-only)", value=str(rec.get("Narrative","")), height=240, key="narrative_readonly_review", disabled=True)
        comments = st.text_area("Reviewer Comments", key="rev_comments_queue_auth")
        c = st.columns(3)
        if can(user,"CanReview"):
            if c[0].button("Approve", key="btn_approve_queue_auth"):
                if not can(user,"CanApprove"):
                    st.error("No permission to approve.")
                else:
                    row = rec; row["Status"] = "Approved"; row["ReviewedBy"] = user.get("Username"); row["ReviewedAt"] = datetime.now().strftime("%Y-%m-%d %H:%M"); row["ReviewerComments"] = comments
                    data["Incidents"] = upsert_row(data["Incidents"], row, key=PRIMARY_KEY)
                    if st.session_state.get("autosave", True): save_to_path(data, file_path)
                    st.success("Approved.")
            if c[1].button("Reject", key="btn_reject_queue_auth"):
                row = rec; row["Status"] = "Rejected"; row["ReviewedBy"] = user.get("Username"); row["ReviewedAt"] = datetime.now().strftime("%Y-%m-%d %H:%M"); row["ReviewerComments"] = comments or "Please revise."
                data["Incidents"] = upsert_row(data["Incidents"], row, key=PRIMARY_KEY)
                if st.session_state.get("autosave", True): save_to_path(data, file_path)
                st.warning("Rejected.")
            if c[2].button("Send back to Draft", key="btn_backtodraft_queue_auth"):
                row = rec; row["Status"] = "Draft"; row["ReviewerComments"] = comments
                data["Incidents"] = upsert_row(data["Incidents"], row, key=PRIMARY_KEY)
                if st.session_state.get("autosave", True): save_to_path(data, file_path)
                st.info("Moved back to Draft.")

with tabs[2]:
    st.header("Rejected Reports")
    if can(user,"CanEditAll"):
        rejected = data["Incidents"][data["Incidents"]["Status"].astype(str) == "Rejected"]
    else:
        rejected = data["Incidents"][(data["Incidents"]["Status"].astype(str) == "Rejected") & (data["Incidents"]["CreatedBy"].astype(str) == user.get("Username"))]
    st.dataframe(rejected, use_container_width=True, hide_index=True, key="grid_rejected_auth")
    selr = None
    if not rejected.empty:
        selr = st.selectbox("Pick a Rejected Incident", options=rejected[PRIMARY_KEY].astype(str).tolist(), index=None, placeholder="Choose...", key="pick_rejected_auth")
    if selr:
        rec = data["Incidents"][data["Incidents"][PRIMARY_KEY].astype(str) == selr].iloc[0].to_dict()
        st.subheader(f"Incident {selr} — Reviewer Comments")
        st.text_area("Reviewer Comments (read-only)", value=str(rec.get("ReviewerComments","")), height=140, key="rejected_comments_readonly", disabled=True)
        st.write("**Narrative (read-only):**")
        st.text_area("Narrative", value=str(rec.get("Narrative","")), height=240, key="rejected_narrative_readonly", disabled=True)
        c = st.columns(2)
        if c[0].button("Move back to Draft to Edit", key="btn_rejected_to_draft"):
            rec["Status"] = "Draft"
            data["Incidents"] = upsert_row(data["Incidents"], rec, key=PRIMARY_KEY)
            if st.session_state.get("autosave", True): save_to_path(data, file_path)
            st.session_state["edit_incident_preselect"] = str(selr)
            st.session_state["force_edit_mode"] = True
            st.success("Moved to Draft. Go to Write Report → Edit to revise and resubmit.")

with tabs[3]:
    st.header("Approved Reports")
    approved = data["Incidents"][data["Incidents"]["Status"].astype(str) == "Approved"]
    st.dataframe(approved, use_container_width=True, hide_index=True, key="grid_approved_auth")
    sela = None
    if not approved.empty:
        sela = st.selectbox("Pick an Approved Incident", options=approved[PRIMARY_KEY].astype(str).tolist(), index=None, placeholder="Choose...", key="pick_approved_auth")
    if sela:
        rec = data["Incidents"][data["Incidents"][PRIMARY_KEY].astype(str) == sela].iloc[0].to_dict()
        st.subheader(f"Incident {sela}")
        st.write(f"**Type:** {rec.get('IncidentType','')}  |  **Priority:** {rec.get('ResponsePriority','')}  |  **Alarm:** {rec.get('AlarmLevel','')}")
        st.write(f"**Date/Time:** {rec.get('IncidentDate','')} {rec.get('IncidentTime','')}")
        st.write(f"**Location:** {rec.get('LocationName','')} — {rec.get('Address','')} {rec.get('City','')} {rec.get('State','')} {rec.get('PostalCode','')}")
        st.write(f"**Shift:** {rec.get('Shift','')}  |  **Reviewed By:** {rec.get('ReviewedBy','')} at {rec.get('ReviewedAt','')}")
        st.write("**Narrative:**")
        
        # Archive controls
        inc_df = data["Incidents"]
        cur_arch = str(rec.get("ArchiveStatus","ACTIVE") or "ACTIVE")
        if cur_arch != "ARCHIVED":
            if bool(user.get("CanApprove", False)) or bool(user.get("CanReview", False)) or bool(user.get("CanManageUsers", False)):
                if st.button("Move to Archive", key="btn_move_to_archive"):
                    data["Incidents"].loc[data["Incidents"][PRIMARY_KEY].astype(str) == str(sela), "ArchiveStatus"] = "ARCHIVED"
                    save_to_path(data, file_path)
                    st.rerun()
                    st.success("Moved to archive.")
        else:
            st.success("This report is archived.")
        st.text_area("Narrative (read-only)", value=str(rec.get("Narrative","")), height=260, key="narrative_readonly_approved", disabled=True)

        ip = ensure_columns(data.get("Incident_Personnel", pd.DataFrame()), CHILD_TABLES["Incident_Personnel"])
        ia = ensure_columns(data.get("Incident_Apparatus", pd.DataFrame()), CHILD_TABLES["Incident_Apparatus"])
        ip_view = ip[ip[PRIMARY_KEY].astype(str) == str(sela)]
        ia_view = ia[ia[PRIMARY_KEY].astype(str) == str(sela)]
        st.markdown(f"**Personnel on Scene ({len(ip_view)}):**")
        if not ip_view.empty:
            show_person_cols = [c for c in ["Name","Role","Hours","RespondedIn"] if c in ip_view.columns]
            st.dataframe(ip_view[show_person_cols], use_container_width=True, hide_index=True, key="grid_approved_personnel")
        else:
            st.write("_None recorded._")
        st.markdown(f"**Apparatus on Scene ({len(ia_view)}):**")
        if not ia_view.empty:
            show_cols = [c for c in ["Unit","UnitType","Role","Actions"] if c in ia_view.columns]
            st.dataframe(ia_view[show_cols], use_container_width=True, hide_index=True, key="grid_approved_apparatus")
        else:
            st.write("_None recorded._")


with tabs[4]:
    st.header("Archive")
    inc = data["Incidents"].copy()
    if "ArchiveStatus" not in inc.columns:
        inc["ArchiveStatus"] = pd.NA
    inc["ArchiveStatus"] = inc["ArchiveStatus"].fillna("ACTIVE").astype(str)
    archived = inc[inc["ArchiveStatus"] == "ARCHIVED"]
    st.dataframe(archived, use_container_width=True, hide_index=True, key="grid_archive")

    sel_arch = None
    if not archived.empty:
        sel_arch = st.selectbox(
            "Pick an Archived Incident",
            options=archived[PRIMARY_KEY].astype(str).tolist(),
            index=None,
            placeholder="Choose...",
            key="pick_archive"
        )

    if sel_arch:
        rec = inc[inc[PRIMARY_KEY].astype(str) == str(sel_arch)].iloc[0].to_dict()
        st.subheader(f"Incident {sel_arch}")
        st.write(f"**Status:** {rec.get('Status','')}  |  **Archived:** {rec.get('ArchiveStatus','')}")
        st.write(f"**Date/Time:** {rec.get('IncidentDate','')} {rec.get('IncidentTime','')}")
        st.write(f"**Type:** {rec.get('IncidentType','')}  |  **Priority:** {rec.get('ResponsePriority','')}  |  **Alarm:** {rec.get('AlarmLevel','')}")
        st.write(f"**Caller:** {rec.get('CallerName','')} ({rec.get('CallerPhone','')})")
        st.write(f"**Location:** {rec.get('LocationName','')} — {rec.get('Address','')}, {rec.get('City','')} {rec.get('State','')} {rec.get('PostalCode','')}")
        st.write("**Narrative:**")
        st.text_area("Narrative (read-only)", value=str(rec.get("Narrative","")), height=260, key="narrative_archive", disabled=True)

        ip = ensure_columns(data.get("Incident_Personnel", pd.DataFrame()), CHILD_TABLES["Incident_Personnel"])
        ia = ensure_columns(data.get("Incident_Apparatus", pd.DataFrame()), CHILD_TABLES["Incident_Apparatus"])
        ip_view = ip[ip[PRIMARY_KEY].astype(str) == str(sel_arch)]
        ia_view = ia[ia[PRIMARY_KEY].astype(str) == str(sel_arch)]

        st.markdown(f"**Personnel on Scene ({len(ip_view)}):**")
        if not ip_view.empty:
            cols = [c for c in ["PersonnelID","Name","Role","Hours","RespondedIn"] if c in ip_view.columns]
            st.dataframe(ip_view[cols], use_container_width=True, hide_index=True, key="grid_archive_personnel")
        else:
            st.write("_None recorded._")

        st.markdown(f"**Apparatus on Scene ({len(ia_view)}):**")
        if not ia_view.empty:
            cols = [c for c in ["ApparatusID","Unit","UnitType","Role","Actions"] if c in ia_view.columns]
            st.dataframe(ia_view[cols], use_container_width=True, hide_index=True, key="grid_archive_apparatus")
        else:
            st.write("_None recorded._")

        # Admin-only delete from archive
        can_delete = bool(user.get("CanDeleteArchive", False))
        if can_delete:
            st.divider()
            st.warning("Delete permanently removes this report and its related personnel/apparatus rows.")
            confirm = st.checkbox("I understand this cannot be undone.", key="confirm_delete_archive")
            if st.button("Delete from Archive (Permanent)", type="primary", disabled=not confirm, key="btn_delete_archive"):
                # Delete from Incidents
                data["Incidents"] = data["Incidents"][data["Incidents"][PRIMARY_KEY].astype(str) != str(sel_arch)].copy()
                # Delete related child rows
                for t in ["Incident_Personnel","Incident_Apparatus","Incident_Times","Incident_Actions"]:
                    if t in data and PRIMARY_KEY in data[t].columns:
                        data[t] = data[t][data[t][PRIMARY_KEY].astype(str) != str(sel_arch)].copy()
                save_to_path(data, file_path)
                st.rerun()
                st.success("Deleted from archive.")
        else:
            st.info("Only Admin users with Delete permission can permanently delete archived reports.")


with tabs[5]:
    st.header("Rosters")
    st.caption("Edit, then click Save. Rank is free text (letters allowed).")
    # Roster editing still permission-gated in earlier build; keep simple here:
    personnel = ensure_columns(data.get("Personnel", pd.DataFrame()), PERSONNEL_SCHEMA)
    if "Rank" in personnel.columns:
        personnel["Rank"] = personnel["Rank"].astype(str)
    personnel_edit = st.data_editor(personnel, num_rows="dynamic", use_container_width=True, key="editor_personnel_auth")
    apparatus = ensure_columns(data.get("Apparatus", pd.DataFrame()), APPARATUS_SCHEMA)
    apparatus_edit = st.data_editor(apparatus, num_rows="dynamic", use_container_width=True, key="editor_apparatus_auth")
    c = st.columns(3)
    if c[0].button("Save Personnel to Excel", key="save_personnel_auth"):
        data["Personnel"] = ensure_columns(personnel_edit, PERSONNEL_SCHEMA)
        ok, err = save_to_path(data, file_path)
        st.success("Saved.") if ok else st.error(err)
    if c[1].button("Save Apparatus to Excel", key="save_apparatus_auth"):
        data["Apparatus"] = ensure_columns(apparatus_edit, APPARATUS_SCHEMA)
        ok, err = save_to_path(data, file_path)
        st.success("Saved.") if ok else st.error(err)

with tabs[6]:
    st.header("Print")
    status = st.selectbox("Filter by Status", options=["","Approved","Submitted","Draft","Rejected"], key="print_status_auth")
    base = data["Incidents"].copy()
    if status: base = base[base["Status"].astype(str) == status]
    st.dataframe(base, use_container_width=True, hide_index=True, key="grid_print_auth")
    sel = None
    if not base.empty:
        sel = st.selectbox("Pick an Incident", options=base[PRIMARY_KEY].astype(str).tolist(), index=None, key="print_pick_auth")
    if sel:
        rec = data["Incidents"][data["Incidents"][PRIMARY_KEY].astype(str) == sel].iloc[0].to_dict()
        st.subheader(f"Incident {sel}")
        st.write(f"**Type:** {rec.get('IncidentType','')}  |  **Priority:** {rec.get('ResponsePriority','')}  |  **Alarm:** {rec.get('AlarmLevel','')}")
        st.write(f"**Location:** {rec.get('Address','')} {rec.get('City','')} {rec.get('State','')} {rec.get('PostalCode','')}")
        st.write("**Narrative:**")
        st.text_area("Narrative (read-only)", value=str(rec.get("Narrative","")), height=220, key="narrative_readonly_print", disabled=True)
        ip = ensure_columns(data.get("Incident_Personnel", pd.DataFrame()), CHILD_TABLES["Incident_Personnel"])
        ia = ensure_columns(data.get("Incident_Apparatus", pd.DataFrame()), CHILD_TABLES["Incident_Apparatus"])
        ip_view = ip[ip[PRIMARY_KEY].astype(str) == str(sel)]
        ia_view = ia[ia[PRIMARY_KEY].astype(str) == str(sel)]
        st.markdown(f"**Personnel on Scene ({len(ip_view)}):**")
        show_person_cols = [c for c in ["Name","Role","Hours","RespondedIn"] if c in ip_view.columns]
        st.dataframe(ip_view[show_person_cols] if not ip_view.empty else ip_view, use_container_width=True, hide_index=True, key="grid_print_personnel")
        st.markdown(f"**Apparatus on Scene ({len(ia_view)}):**")
        show_cols = [c for c in ["Unit","UnitType","Role","Actions"] if c in ia_view.columns]
        st.dataframe(ia_view[show_cols] if not ia_view.empty else ia_view, use_container_width=True, hide_index=True, key="grid_print_apparatus")

        # --- PRINT / EXPORT CONTROLS (Print tab only) ---
        import streamlit.components.v1 as components
        import html as _html
        import io
        try:
            from reportlab.lib.pagesizes import LETTER
            from reportlab.pdfgen import canvas
            from reportlab.lib.units import inch
            _PDF_OK = True
        except Exception:
            _PDF_OK = False

        # Resolve selected incident record
        try:
            rec = base[base[PRIMARY_KEY].astype(str) == str(sel)].iloc[0].to_dict()
        except Exception:
            rec = {}

        # Times
        times_df = ensure_columns(data.get("Incident_Times", pd.DataFrame()), CHILD_TABLES["Incident_Times"])
        trow = {}
        if not times_df.empty:
            _m = times_df[PRIMARY_KEY].astype(str) == str(sel)
            if _m.any():
                trow = times_df[_m].iloc[0].to_dict()

        # Personnel/Apparatus for this incident (fresh views)
        ip_df = ensure_columns(data.get("Incident_Personnel", pd.DataFrame()), CHILD_TABLES["Incident_Personnel"])
        ia_df = ensure_columns(data.get("Incident_Apparatus", pd.DataFrame()), CHILD_TABLES["Incident_Apparatus"])
        ip_view2 = ip_df[ip_df[PRIMARY_KEY].astype(str) == str(sel)]
        ia_view2 = ia_df[ia_df[PRIMARY_KEY].astype(str) == str(sel)]

        def esc(x): return _html.escape("" if x is None else str(x))

        html_report = f"""
<h2>Incident #{esc(sel)}</h2>
<b>Date/Time:</b> {esc(rec.get('IncidentDate',''))} {esc(rec.get('IncidentTime',''))}<br>
<b>Location:</b> {esc(rec.get('LocationName',''))} — {esc(rec.get('Address',''))} {esc(rec.get('City',''))} {esc(rec.get('State',''))} {esc(rec.get('PostalCode',''))}<br>
<b>Caller:</b> {esc(rec.get('CallerName','') or 'N/A')} ({esc(rec.get('CallerPhone','') or 'N/A')})<br>
<b>Report Writer:</b> {esc(rec.get('ReportWriter','') or rec.get('CreatedBy','') or 'N/A')} &nbsp;&nbsp; <b>Approver:</b> {esc(rec.get('Approver','') or rec.get('ReviewedBy','') or 'N/A')}<br>
<b>Times:</b> Alarm {esc(trow.get('Alarm',''))} | Enroute {esc(trow.get('Enroute',''))} | Arrival {esc(trow.get('Arrival',''))} | Clear {esc(trow.get('Clear',''))}<br><br>
<h3>Narrative</h3>
<div style="white-space: pre-wrap;">{esc(rec.get('Narrative',''))}</div>
<br>
<h3>Personnel on Scene</h3>
{ip_view2.to_html(index=False)}
<br>
<h3>Apparatus on Scene</h3>
{ia_view2.to_html(index=False)}
"""

        c1, c2, c3 = st.columns(3)

        # 1) Print (browser dialog)
        if c1.button("🖨️ Print Page", key=f"print_tab_print_{sel}"):
            components.html("<script>window.print()</script>", height=0, width=0)

        # 2) Download HTML (works everywhere)
        c2.download_button("⬇️ Download HTML", html_report,
                           file_name=f"Incident_{sel}.html", mime="text/html",
                           key=f"print_tab_html_{sel}")

        # 3) Optional PDF (requires 'reportlab' in requirements; otherwise this button won't show)
        if _PDF_OK and c3.button("📄 Download PDF", key=f"print_tab_pdf_{sel}"):
            try:
                buf = io.BytesIO()
                c = canvas.Canvas(buf, pagesize=LETTER)
                text = c.beginText(0.5*inch, 10.5*inch)
                text.setFont("Helvetica", 10)
                raw = (html_report.replace("<br>", "\n")
                                  .replace("<h3>", "\n")
                                  .replace("</h3>", "")
                                  .replace("<div", "")
                                  .replace("</div>", ""))
                for line in raw.split("\n"):
                    text.textLine(line)
                c.drawText(text); c.showPage(); c.save()
                buf.seek(0)
                st.download_button("Save PDF", data=buf,
                                   file_name=f"Incident_{sel}.pdf", mime="application/pdf",
                                   key=f"print_tab_pdf_dl_{sel}")
            except Exception as e:
                st.error(f"PDF failed: {e}")


with tabs[7]:
    st.header("Export")
    if st.button("Build Excel for Download", key="btn_build_export_auth"):
        payload = save_workbook_to_bytes(data)
        st.download_button("Download Excel", data=payload, file_name="fire_incident_db_export.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key="download_export_auth")
    if st.button("Overwrite Source File Now", key="btn_overwrite_source_auth"):
        ok, err = save_to_path(data, file_path)
        if ok: st.success(f"Wrote: {file_path}")
        else: st.error(f"Failed: {err}")

with tabs[8]:
    st.header("Admin — User Management & Permissions")
    users_df = apply_role_presets(ensure_columns(data.get("Users", pd.DataFrame()), USERS_SCHEMA))
    users_edit = st.data_editor(users_df, num_rows="dynamic", use_container_width=True, key="editor_users_auth")
    c = st.columns(3)
    if c[0].button("Save Users to Excel", key="save_users_auth"):
        users_edit = ensure_columns(users_edit, USERS_SCHEMA)
        ok, err = save_to_path({**data, "Users": users_edit}, file_path)
        if ok:
            data["Users"] = users_edit
            st.success("Users saved.")
        else:
            st.error(err)

with tabs[9]:
    st.header("Diagnostics")
    st.write(f"**App dir:** {os.path.dirname(__file__)}")
    st.write(f"**Excel path:** {file_path}  |  Exists: {'✅' if os.path.exists(file_path) else '❌'}")
    try:
        xls = pd.ExcelFile(file_path); st.write("**Sheets:**", xls.sheet_names)
    except Exception as e:
        st.error(f"Open failed: {e}")
    st.write("**Personnel Top 10:**")
    st.dataframe(data['Personnel'].head(10), use_container_width=True)
    st.write("**Apparatus Top 10:**")
    st.dataframe(data['Apparatus'].head(10), use_container_width=True)
    st.write("**Users Top 10:**")
    st.dataframe(data['Users'].head(10), use_container_width=True)