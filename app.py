
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

def load_workbook(path: str) -> Dict[str, pd.DataFrame]:
    try:
        xls = pd.ExcelFile(path)
        return {name.strip(): xls.parse(name) for name in xls.sheet_names}
    except Exception as e:
        st.error(f"Failed to load: {e}")
        return {}

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

def label_rank_first_last(row: pd.Series) -> str:
    fn = str(row.get("FirstName") or "").strip()
    ln = str(row.get("LastName") or "").strip()
    rk = str(row.get("Rank") or "").strip()
    parts = [p for p in [rk, fn, ln] if p]
    return " ".join(parts).strip()

def build_names(df: pd.DataFrame, active_only: bool) -> list:
    if "Name" in df and df["Name"].notna().any():
        s = df["Name"].astype(str)
    elif "FullName" in df and df["FullName"].notna().any():
        s = df["FullName"].astype(str)
    elif all(c in df.columns for c in ["FirstName","LastName","Rank"]):
        s = df.apply(label_rank_first_last, axis=1)
    elif all(c in df.columns for c in ["FirstName","LastName"]):
        s = (df["FirstName"].fillna("").astype(str).str.strip() + " " + df["LastName"].fillna("").astype(str).str.strip()).str.strip()
    else:
        s = pd.Series([], dtype=str)
    if active_only and "Active" in df.columns:
        mask = df["Active"].astype(str).str.strip().str.lower().isin(["yes","true","1","active","y"])
        s = s[mask]
    vals = s.dropna().map(lambda x: x.strip()).replace("", pd.NA).dropna().unique().tolist()
    return sorted(set(vals))

def build_units(df: pd.DataFrame, active_only: bool) -> list:
    for col in ["UnitNumber","CallSign","Name"]:
        if col in df.columns and df[col].notna().any():
            s = df[col].astype(str); break
    else:
        s = pd.Series([], dtype=str)
    if active_only and "Active" in df.columns:
        mask = df["Active"].astype(str).str.strip().str.lower().isin(["yes","true","1","active","y"])
        s = s[mask]
    vals = s.dropna().map(lambda x: x.strip()).replace("", pd.NA).dropna().unique().tolist()
    return sorted(set(vals))

# ---------------- App Boot ----------------
st.sidebar.title("📝 Fire Incident Reports")
file_path = st.sidebar.text_input("Excel path", value=DEFAULT_FILE, key="sidebar_path")
st.sidebar.write(f"**File exists:** {'✅' if os.path.exists(file_path) else '❌'}")
uploaded = st.sidebar.file_uploader("Upload/replace workbook (.xlsx)", type=["xlsx"], key="sidebar_upload")
if uploaded:
    with open(file_path, "wb") as f: f.write(uploaded.read())
    st.sidebar.success(f"Saved to {file_path}")
st.session_state.setdefault("autosave", True)
st.session_state["autosave"] = st.sidebar.toggle("Autosave to Excel", value=True)
st.session_state.setdefault("filter_active_only", False)  # default OFF
st.session_state["filter_active_only"] = st.sidebar.toggle("Show ACTIVE roster only", value=st.session_state["filter_active_only"])
st.session_state.setdefault("roster_source", "excel")  # default to excel for clarity
st.session_state["roster_source"] = st.sidebar.selectbox("Roster source", options=["excel","app"], index=0)

data: Dict[str, pd.DataFrame] = {}
if os.path.exists(file_path): data = load_workbook(file_path)
if not data: st.info("Upload or point to your Excel workbook to begin."); st.stop()

# Ensure base tables
ensure_table(data, "Incidents", [
    PRIMARY_KEY,"IncidentDate","IncidentTime","IncidentType","ResponsePriority","AlarmLevel","Shift",
    "LocationName","Address","City","State","PostalCode","Latitude","Longitude",
    "Narrative","Status","CreatedBy","SubmittedAt","ReviewedBy","ReviewedAt","ReviewerComments"
])
ensure_table(data, "Personnel", PERSONNEL_SCHEMA)
ensure_table(data, "Apparatus", APPARATUS_SCHEMA)
for t, cols in CHILD_TABLES.items(): ensure_table(data, t, cols)

# Session rosters
st.session_state.setdefault("roster_personnel", data["Personnel"].copy())
st.session_state.setdefault("roster_apparatus", data["Apparatus"].copy())

lookups = get_lookups(data)

def get_roster(table: str) -> pd.DataFrame:
    if st.session_state.get("roster_source","excel") == "app":
        return st.session_state["roster_personnel"].copy() if table == "Personnel" else st.session_state["roster_apparatus"].copy()
    return data[table].copy()

tabs = st.tabs(["Write Report","Review Queue","Rosters","Print","Export","Diagnostics"])

with tabs[0]:
    st.header("Write Report")
    master = data["Incidents"]
    mode = st.radio("Mode", ["New","Edit"], horizontal=True, key="mode_write")
    defaults = {}; selected = None
    if mode == "Edit":
        options_df = master
        options = options_df[PRIMARY_KEY].dropna().astype(str).tolist() if PRIMARY_KEY in options_df.columns else []
        selected = st.selectbox("Select IncidentNumber", options=options, index=None, placeholder="Choose...", key="pick_edit_write")
        if selected:
            defaults = master[master[PRIMARY_KEY].astype(str) == selected].iloc[0].to_dict()

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
    state = c10.text_input("State", value=str(defaults.get("State","")) if defaults else "", key="w_state")
    postal = c11.text_input("PostalCode", value=str(defaults.get("PostalCode","")) if defaults else "", key="w_postal")
    shift = c12.text_input("Shift", value=str(defaults.get("Shift","")) if defaults else "", key="w_shift")

    st.subheader("Narrative")
    narrative = st.text_area("Write full narrative here", value=str(defaults.get("Narrative","")) if defaults else "", height=300, key="w_narrative")

    st.subheader("All Members on Scene")
    roster_people = ensure_columns(get_roster("Personnel"), PERSONNEL_SCHEMA)
    names = build_names(roster_people, active_only=st.session_state.get("filter_active_only", False))
    st.caption(f"Detected **{len(names)}** member options from roster source: {st.session_state['roster_source']}")
    picked_people = st.multiselect("Pick members", options=names, key="w_pick_people")

    st.subheader("Apparatus on Scene")
    roster_units = ensure_columns(get_roster("Apparatus"), APPARATUS_SCHEMA)
    units = build_units(roster_units, active_only=st.session_state.get("filter_active_only", False))
    st.caption(f"Detected **{len(units)}** apparatus options from roster source: {st.session_state['roster_source']}")
    picked_units = st.multiselect("Pick apparatus units", options=units, key="w_pick_units")

with tabs[5]:
    st.header("Diagnostics")
    st.write("### Paths and Files")
    st.write(f"**Current working directory:** {os.getcwd()}")
    st.write(f"**App directory files:** {os.listdir(os.path.dirname(__file__))}")
    st.write(f"**Excel path set to:** {file_path}")
    st.write(f"**Excel exists?** {'✅ Yes' if os.path.exists(file_path) else '❌ No'}")

    st.write("### Workbook Overview")
    if os.path.exists(file_path):
        try:
            xls = pd.ExcelFile(file_path)
            st.write("**Sheet names:**", xls.sheet_names)
        except Exception as e:
            st.error(f"Could not open workbook: {e}")

    st.write("### Personnel sheet (top 10)")
    st.dataframe(roster_people.head(10), use_container_width=True)
    st.write(f"**Computed member options ({len(names)}):**", names[:50])

    st.write("### Apparatus sheet (top 10)")
    st.dataframe(roster_units.head(10), use_container_width=True)
    st.write(f"**Computed unit options ({len(units)}):**", units[:50])
