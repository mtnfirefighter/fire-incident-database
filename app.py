import streamlit as st
import pandas as pd
import io
from typing import Dict, Tuple, List
from datetime import datetime

st.set_page_config(page_title="Fire Incident DB", page_icon="🚒", layout="wide")

# --- Settings ---
import os
FILE_PATH = os.path.join(os.path.dirname(__file__), "fire_incident_db.xlsx")
LOOKUP_PREFIX = "List_"

# --- Helpers ---
def load_workbook(path: str) -> Dict[str, pd.DataFrame]:
    try:
        xls = pd.ExcelFile(path)
        data = {name: xls.parse(name) for name in xls.sheet_names}
        return data
    except Exception as e:
        st.error(f"Failed to load Excel file: {e}")
        return {}

def ensure_columns(df: pd.DataFrame, needed: List[str]) -> pd.DataFrame:
    for c in needed:
        if c not in df.columns:
            df[c] = pd.NA
    return df

def init_state():
    if "data" not in st.session_state:
        st.session_state.data = {}
    if "file_path" not in st.session_state:
        st.session_state.file_path = FILE_PATH

def save_workbook_to_buffer(dfs: Dict[str, pd.DataFrame]) -> bytes:
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="xlsxwriter") as writer:
        for sheet, df in dfs.items():
            df_export = df.copy()
            # ensure index not saved
            df_export.to_excel(writer, sheet_name=sheet, index=False)
    buf.seek(0)
    return buf.read()

def incident_form(defaults: dict, lookups: Dict[str, pd.DataFrame]) -> dict:
    left, right = st.columns(2)
    with left:
        incident_id = st.number_input("IncidentID", value=int(defaults.get("IncidentID", 0)) if pd.notna(defaults.get("IncidentID", pd.NA)) else 0, min_value=0, step=1)
        date = st.date_input("Date", value=defaults.get("Date", pd.Timestamp.today()))
        incident_type = st.selectbox("IncidentType", options=lookups.get("List_IncidentTypes", pd.DataFrame(columns=["IncidentTypes"]))["IncidentTypes"].dropna().tolist() if "List_IncidentTypes" in lookups else [], index=None, placeholder="Select...")
        priority = st.selectbox("Priority", options=lookups.get("List_Priorities", pd.DataFrame(columns=["Priorities"]))["Priorities"].dropna().tolist() if "List_Priorities" in lookups else [], index=None, placeholder="Select...")
        disposition = st.selectbox("Disposition", options=lookups.get("List_Dispositions", pd.DataFrame(columns=["Dispositions"]))["Dispositions"].dropna().tolist() if "List_Dispositions" in lookups else [], index=None, placeholder="Select...")
    with right:
        address = st.text_input("Address", value=str(defaults.get("Address", "")) if pd.notna(defaults.get("Address", pd.NA)) else "")
        city = st.text_input("City", value=str(defaults.get("City", "")) if pd.notna(defaults.get("City", pd.NA)) else "")
        state = st.selectbox("State", options=lookups.get("List_States", pd.DataFrame(columns=["States"]))["States"].dropna().tolist() if "List_States" in lookups else [], index=None, placeholder="Select...")
        postal = st.text_input("PostalCode", value=str(defaults.get("PostalCode", "")) if pd.notna(defaults.get("PostalCode", pd.NA)) else "")
        cross = st.text_input("CrossStreets", value=str(defaults.get("CrossStreets", "")) if pd.notna(defaults.get("CrossStreets", pd.NA)) else "")
    return {
        "IncidentID": int(incident_id),
        "Date": pd.to_datetime(date),
        "IncidentType": incident_type,
        "Priority": priority,
        "Disposition": disposition,
        "Address": address,
        "City": city,
        "State": state,
        "PostalCode": postal,
        "CrossStreets": cross,
    }

def related_detail_ui(name: str, df: pd.DataFrame, incident_id: int, lookup_df: pd.DataFrame | None, label_col: str, cols: List[str]) -> pd.DataFrame:
    st.subheader(name.replace("_", " "))
    filtered = df[df["IncidentID"] == incident_id].copy() if "IncidentID" in df.columns else df.copy()
    st.dataframe(filtered, use_container_width=True, hide_index=True)
    with st.expander(f"Add to {name}"):
        add_cols = st.columns(min(4, len(cols)))
        inputs = {}
        for i, col in enumerate(cols):
            with add_cols[i % len(add_cols)]:
                if lookup_df is not None and col == label_col and label_col in lookup_df.columns:
                    options = lookup_df[label_col].dropna().tolist()
                    inputs[col] = st.selectbox(col, options=options, index=None, placeholder="Select...")
                else:
                    if col.lower().endswith("time") or col in ["Alarm", "Arrival", "Clear"]:
                        inputs[col] = st.text_input(col, placeholder="HH:MM")
                    else:
                        inputs[col] = st.text_input(col)
        if st.button(f"Add row to {name}", key=f"add_{name}"):
            new_row = {c: inputs.get(c, "") for c in cols}
            new_row["IncidentID"] = incident_id
            df = pd.concat([df, pd.DataFrame([new_row])], ignore_index=True)
            st.success("Added.")
    return df

# --- App ---
init_state()

st.sidebar.title("🚒 Fire Incident DB")
choice = st.sidebar.radio("Go to", ["Load/Replace Workbook", "Browse & Filter", "Add / Edit Incident", "Related Records", "Export / Save"])

# Load section
with st.sidebar.expander("Settings"):
    st.session_state.file_path = st.text_input("Excel file path (in app directory)", value=st.session_state.file_path)

if choice == "Load/Replace Workbook":
    st.title("Load / Replace Workbook")
    uploaded = st.file_uploader("Upload Excel file (.xlsx)", type=["xlsx"])
    if uploaded:
        with open(st.session_state.file_path, "wb") as f:
            f.write(uploaded.read())
        st.success(f"Saved to {st.session_state.file_path}")
    st.info("After uploading, navigate to another tab to work with the data.")

# Load data if available
data = load_workbook(st.session_state.file_path) if os.path.exists(st.session_state.file_path) else {}
if data:
    # Normalize expected core sheets
    core_tables = ["Incidents","Incident_Detail","Incident_Times","Incident_Personnel","Incident_Apparatus","Incident_Actions"]
    for t in core_tables:
        data[t] = data.get(t, pd.DataFrame())
    # Ensure expected columns for Incidents
    needed_cols = ["IncidentID","Date","IncidentType","Priority","Disposition","Address","City","State","PostalCode","CrossStreets"]
    data["Incidents"] = ensure_columns(data["Incidents"], needed_cols)

# Lookup tables
lookups = {k: v for k, v in data.items() if k.startswith(LOOKUP_PREFIX)}
personnel_lu = data.get("Personnel", pd.DataFrame())
apparatus_lu = data.get("Apparatus", pd.DataFrame())
actions_lu = data.get("List_Actions", pd.DataFrame())

if choice == "Browse & Filter":
    st.title("Browse & Filter Incidents")
    if not data:
        st.warning("Load a workbook first.")
    else:
        df = data["Incidents"].copy()
        # Filters
        c1, c2, c3, c4 = st.columns(4)
        with c1:
            incident_type = st.selectbox("IncidentType", options=lookups.get("List_IncidentTypes", pd.DataFrame(columns=["IncidentTypes"]))["IncidentTypes"].dropna().tolist() if "List_IncidentTypes" in lookups else [], index=None, placeholder="All")
        with c2:
            priority = st.selectbox("Priority", options=lookups.get("List_Priorities", pd.DataFrame(columns=["Priorities"]))["Priorities"].dropna().tolist() if "List_Priorities" in lookups else [], index=None, placeholder="All")
        with c3:
            city = st.text_input("City contains")
        with c4:
            date_range = st.date_input("Date range", value=[])
        if incident_type:
            df = df[df["IncidentType"] == incident_type]
        if priority:
            df = df[df["Priority"] == priority]
        if city:
            df = df[df["City"].astype(str).str.contains(city, case=False, na=False)]
        if isinstance(date_range, list) and len(date_range) == 2:
            start, end = pd.to_datetime(date_range[0]), pd.to_datetime(date_range[1])
            df = df[(pd.to_datetime(df["Date"], errors="coerce") >= start) & (pd.to_datetime(df["Date"], errors="coerce") <= end)]
        st.dataframe(df, use_container_width=True, hide_index=True)

if choice == "Add / Edit Incident":
    st.title("Add / Edit Incident")
    if not data:
        st.warning("Load a workbook first.")
    else:
        df = data["Incidents"]
        mode = st.radio("Mode", ["Add", "Edit"], horizontal=True)
        if mode == "Edit" and not df.empty:
            ids = df["IncidentID"].dropna().astype(int).tolist()
            selected = st.selectbox("Select IncidentID", options=ids, index=None, placeholder="Choose...")
            defaults = df[df["IncidentID"] == selected].iloc[0].to_dict() if selected is not None and (df["IncidentID"] == selected).any() else {}
        else:
            defaults = {}
        form_vals = incident_form(defaults, lookups)
        if st.button("Save Incident"):
            # Insert or update
            if mode == "Edit" and defaults:
                idx = df.index[df["IncidentID"] == defaults["IncidentID"]]
                for k, v in form_vals.items():
                    df.loc[idx, k] = v
                st.success("Incident updated.")
            else:
                data["Incidents"] = pd.concat([df, pd.DataFrame([form_vals])], ignore_index=True)
                st.success("Incident added.")

if choice == "Related Records":
    st.title("Related Records")
    if not data or data["Incidents"].empty:
        st.warning("Load a workbook and ensure there is at least one incident.")
    else:
        ids = data["Incidents"]["IncidentID"].dropna().astype(int).tolist()
        incident_id = st.selectbox("Choose IncidentID", options=ids, index=None, placeholder="Select...")
        if incident_id is not None:
            # Times
            data["Incident_Times"] = ensure_columns(data["Incident_Times"], ["IncidentID","Alarm","Arrival","Clear"])
            data["Incident_Times"] = related_detail_ui("Incident_Times", data["Incident_Times"], incident_id, None, "", ["Alarm","Arrival","Clear"])

            # Personnel
            data["Incident_Personnel"] = ensure_columns(data["Incident_Personnel"], ["IncidentID","Name","Role"])
            data["Incident_Personnel"] = related_detail_ui("Incident_Personnel", data["Incident_Personnel"], incident_id, personnel_lu if not personnel_lu.empty else None, "Name", ["Name","Role"])

            # Apparatus
            label = "ApparatusID" if "ApparatusID" in data.get("Incident_Access", pd.DataFrame()).columns else "Unit"
            data["Incident_Aparatus"] = data.get("Incident_Apparatus", pd.DataFrame())
            data["Incident_Apparatus"] = ensure_columns(data["Incident_Apparatus"], ["IncidentID","Unit","Action"])
            data["Incident_Apparatus"] = related_detail_ui("Incident_Apparatus", data["Incident_Apparatus"], incident_id, apparatus_lu if not apparatus_lu.empty else None, "Unit", ["Unit","Action"])

            # Actions
            data["Incident_Actions"] = ensure_columns(data["Incident_Actions"], ["IncidentID","Action"])
            data["Incident_Actions"] = related_detail_ui("Incident_Actions", data["Incident_Actions"], incident_id, actions_lu if not actions_lu.empty else None, "Actions", ["Action"])

if choice == "Export / Save":
    st.title("Export / Save")
    if not data:
        st.warning("Load a workbook first.")
    else:
        st.write("Download a fresh Excel with current data frames (this will not overwrite the source file unless you choose to).")
        if st.button("Build Excel for Download"):
            payload = save_workbook_to_buffer(data)
            st.download_button("Download Excel", data=payload, file_name="fire_incident_db_export.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        st.divider()
        if st.button("Overwrite Source File"):
            payload = save_workbook_to_buffer(data)
            with open(st.session_state.file_path, "wb") as f:
                f.write(payload)
            st.success(f"Overwrote {st.session_state.file_path}")
