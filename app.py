
import os
import io
from datetime import datetime, date
from typing import Dict, List

import pandas as pd
import streamlit as st

st.set_page_config(page_title="Fire Incident DB", page_icon="🚒", layout="wide")

# ---------- Settings ----------
DEFAULT_FILE = os.path.join(os.path.dirname(__file__), "fire_incident_db.xlsx")
PRIMARY_KEY = "IncidentNumber"  # aligns with your workbook
CHILD_TABLES = {
    "Incident_Times": ["IncidentNumber","Alarm","Enroute","Arrival","Clear"],
    "Incident_Personnel": ["IncidentNumber","Name","Role","Hours"],
    "Incident_Apparatus": ["IncidentNumber","Unit","UnitType","Role","Actions"],
    "Incident_Actions": ["IncidentNumber","Action","Notes"],
}
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

def render_dynamic_form(df: pd.DataFrame, lookups: Dict[str, List[str]], defaults: dict) -> dict:
    """Render inputs for each column in Incidents based on type/name hints."""
    vals = {}
    prefer_first = [PRIMARY_KEY, "IncidentDate", "IncidentTime", "IncidentType", "ResponsePriority", "AlarmLevel",
                    "Shift","LocationName","Address","City","State","PostalCode",
                    "Latitude","Longitude"]
    columns = [c for c in prefer_first if c in df.columns] + [c for c in df.columns if c not in prefer_first]
    cols3 = st.columns(3)
    for i, col in enumerate(columns):
        with cols3[i % 3]:
            current = defaults.get(col, None)
            if col in lookups:
                options = lookups[col]
                idx = options.index(current) if isinstance(current, str) and current in options else None
                vals[col] = st.selectbox(col, options=options, index=idx, placeholder=f"Select {col}...")
            elif col in DATE_LIKE:
                d = pd.to_datetime(current).date() if pd.notna(current) else date.today()
                vals[col] = st.date_input(col, value=d)
            elif col in TIME_LIKE:
                vals[col] = st.text_input(col, value=str(current) if pd.notna(current) else "", placeholder="HH:MM")
            else:
                if col in df.select_dtypes(include="number").columns:
                    base = float(current) if pd.notna(current) else 0.0
                    vals[col] = st.number_input(col, value=base)
                else:
                    vals[col] = st.text_input(col, value=str(current) if pd.notna(current) else "")
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
        # Ensure all columns exist
        for k in row.keys():
            if k not in df.columns:
                df[k] = pd.NA
        df = pd.concat([df, pd.DataFrame([row])], ignore_index=True)
    return df

def related_editor(table_name: str, data: Dict[str, pd.DataFrame], lookups: Dict[str, List[str]], incident_number: str):
    st.subheader(table_name.replace("_", " "))
    df = data.get(table_name, pd.DataFrame())
    if df.empty and table_name in CHILD_TABLES:
        df = pd.DataFrame(columns=CHILD_TABLES[table_name])
    if PRIMARY_KEY not in df.columns:
        df[PRIMARY_KEY] = pd.NA
    view = df[df[PRIMARY_KEY].astype(str) == str(incident_number)].copy()
    st.dataframe(view, use_container_width=True, hide_index=True)
    with st.expander(f"Add to {table_name}"):
        add_vals = {}
        cols = [c for c in df.columns if c != PRIMARY_KEY]
        cols2 = st.columns(3)
        for i, c in enumerate(cols):
            with cols2[i % 3]:
                if c in lookups:
                    opts = lookups[c]
                    add_vals[c] = st.selectbox(c, options=opts, index=None, placeholder=f"Select {c}...")
                elif c in TIME_LIKE:
                    add_vals[c] = st.text_input(c, placeholder="HH:MM")
                else:
                    add_vals[c] = st.text_input(c)
        if st.button(f"Add row to {table_name}"):
            new_row = {PRIMARY_KEY: incident_number}
            new_row.update(add_vals)
            data[table_name] = pd.concat([df, pd.DataFrame([new_row])], ignore_index=True)
            st.success("Added.")

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

    times = data.get("Incident_Times", pd.DataFrame())
    if not times.empty and set(["Alarm","Arrival"]).issubset(times.columns):
        st.subheader("Response Time (Arrival - Alarm) — naive HH:MM diff")
        def to_minutes(s):
            try:
                hh, mm = map(int, str(s).split(":"))
                return hh*60 + mm
            except Exception:
                return None
        tmp = times[[PRIMARY_KEY,"Alarm","Arrival"]].copy()
        tmp["t_alarm"] = tmp["Alarm"].apply(to_minutes)
        tmp["t_arrival"] = tmp["Arrival"].apply(to_minutes)
        tmp["resp_min"] = tmp.apply(lambda r: r["t_arrival"]-r["t_alarm"] if r["t_alarm"] is not None and r["t_arrival"] is not None else None, axis=1)
        st.dataframe(tmp[[PRIMARY_KEY,"Alarm","Arrival","resp_min"]], use_container_width=True, hide_index=True)

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

    for t in ["Incident_Times","Incident_Personnel","Incident_Apparatus","Incident_Actions"]:
        df = data.get(t, pd.DataFrame())
        if not df.empty and PRIMARY_KEY in df.columns:
            view = df[df[PRIMARY_KEY].astype(str) == str(incident_number)]
            if not view.empty:
                st.subheader(t.replace("_"," "))
                st.dataframe(view, use_container_width=True, hide_index=True)

# ---------- App ----------
st.sidebar.title("🚒 Fire Incident DB")
file_path = st.sidebar.text_input("Excel path", value=DEFAULT_FILE)
uploaded = st.sidebar.file_uploader("Upload/replace workbook (.xlsx)", type=["xlsx"])
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

if "Incidents" not in data:
    data["Incidents"] = pd.DataFrame(columns=[PRIMARY_KEY])
for t, cols in CHILD_TABLES.items():
    if t not in data:
        data[t] = pd.DataFrame(columns=cols)

lookups = get_lookups(data)

tab1, tab2, tab3, tab4, tab5 = st.tabs(["Browse","Add / Edit","Related","Reports","Export"])

with tab1:
    st.header("Browse & Filter Incidents")
    base = data["Incidents"].copy()
    fc1, fc2, fc3, fc4 = st.columns(4)
    with fc1:
        val_type = st.selectbox("IncidentType", options=[""]+lookups.get("IncidentType", []))
        if val_type:
            base = base[base.get("IncidentType","").astype(str) == val_type]
    with fc2:
        val_prio = st.selectbox("ResponsePriority", options=[""]+lookups.get("ResponsePriority", []))
        if val_prio:
            base = base[base.get("ResponsePriority","").astype(str) == val_prio]
    with fc3:
        city = st.text_input("City contains")
        if city:
            base = base[base.get("City","").astype(str).str.contains(city, case=False, na=False)]
    with fc4:
        dr = st.date_input("Date range", [])
        if isinstance(dr, list) and len(dr) == 2:
            start, end = pd.to_datetime(dr[0]), pd.to_datetime(dr[1])
            base = base[(pd.to_datetime(base.get("IncidentDate", pd.NaT), errors="coerce") >= start) &
                        (pd.to_datetime(base.get("IncidentDate", pd.NaT), errors="coerce") <= end)]
    st.dataframe(base, use_container_width=True, hide_index=True)

with tab2:
    st.header("Add / Edit Incident")
    master = data["Incidents"]
    mode = st.radio("Mode", ["Add","Edit"], horizontal=True)
    defaults = {}
    if mode == "Edit" and not master.empty and PRIMARY_KEY in master.columns:
        options = master[PRIMARY_KEY].dropna().astype(str).tolist()
        selected = st.selectbox("Select IncidentNumber", options=options, index=None, placeholder="Choose...")
        if selected:
            defaults = master[master[PRIMARY_KEY].astype(str) == selected].iloc[0].to_dict()

    vals = render_dynamic_form(master, lookups, defaults)
    if "IncidentDate" in vals:
        vals["IncidentDate"] = pd.to_datetime(vals["IncidentDate"])
    if st.button("Save Incident"):
        data["Incidents"] = upsert_row(master, vals, key=PRIMARY_KEY)
        st.success("Saved.")

with tab3:
    st.header("Related Records")
    if data["Incidents"].empty:
        st.info("Add an incident first.")
    else:
        inc_id = st.selectbox("IncidentNumber", options=data["Incidents"][PRIMARY_KEY].dropna().astype(str).tolist(), index=None)
        if inc_id:
            for t in ["Incident_Times","Incident_Personnel","Incident_Apparatus","Incident_Actions"]:
                related_editor(t, data, lookups, inc_id)
            st.divider()
            with st.expander("Printable Incident Report"):
                printable_incident(data, inc_id)

with tab4:
    analytics_reports(data)

with tab5:
    st.header("Export")
    if st.button("Build Excel for Download"):
        payload = save_workbook_to_bytes(data)
        st.download_button("Download Excel", data=payload, file_name="fire_incident_db_export.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    if st.button("Overwrite Source File"):
        payload = save_workbook_to_bytes(data)
        with open(file_path, "wb") as f:
            f.write(payload)
        st.success(f"Overwrote: {file_path}")
