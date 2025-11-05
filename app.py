#!/usr/bin/env python3
# Unified Fire Incident Reports App — v4.3.2 (single file)
# Combines your current app plus the print + PDF + apparatus patches.
# - CallSign-first apparatus picker
# - Print tab: full details + Print button + Convert to PDF / HTML fallback
# - Keeps your sign-in, roles, rosters, review/approve flows intact

import os, io
from datetime import datetime, date
from typing import Dict, List, Optional
import pandas as pd
import streamlit as st
import streamlit.components.v1 as components

st.set_page_config(page_title="Fire Incident Reports", page_icon="📝", layout="wide")

DEFAULT_FILE = os.path.join(os.path.dirname(__file__), "fire_incident_db.xlsx")
PRIMARY_KEY = "IncidentNumber"

# ---------- Schemas & Lookups ----------
CHILD_TABLES = {
    "Incident_Times": ["IncidentNumber","Alarm","Enroute","Arrival","Clear"],
    "Incident_Personnel": ["IncidentNumber","Name","Role","Hours","RespondedIn"],
    "Incident_Apparatus": ["IncidentNumber","Unit","UnitType","Role","Actions"],
    "Incident_Actions": ["IncidentNumber","Action","Notes"],
}
PERSONNEL_SCHEMA = ["PersonnelID","Name","UnitNumber","Rank","Badge","Phone","Email","Address","City","State","PostalCode","Certifications","Active","FirstName","LastName","FullName"]
APPARATUS_SCHEMA = ["ApparatusID","UnitNumber","CallSign","UnitType","GPM","TankSize","SeatingCapacity","Station","Active","Name"]
USERS_SCHEMA = ["Username","Password","Role","FullName","Active",
                "CanWrite","CanEditOwn","CanEditAll","CanReview","CanApprove","CanManageUsers","CanEditRosters","CanPrint"]

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
    "Admin":   {"CanWrite":True,"CanEditOwn":True,"CanEditAll":True,"CanReview":True,"CanApprove":True,"CanManageUsers":True,"CanEditRosters":True,"CanPrint":True},
    "Reviewer":{"CanWrite":False,"CanEditOwn":False,"CanEditAll":False,"CanReview":True,"CanApprove":True,"CanManageUsers":False,"CanEditRosters":False,"CanPrint":True},
    "Member":  {"CanWrite":True,"CanEditOwn":True,"CanEditAll":False,"CanReview":False,"CanApprove":False,"CanManageUsers":False,"CanEditRosters":False,"CanPrint":True},
}

# ---------- IO helpers ----------
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

# ---------- Table utilities ----------
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

# ---------- Roster helpers (personnel + apparatus) ----------
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

# CallSign-first (case-insensitive) for apparatus options, falls back to common alternates
def build_unit_options(df: pd.DataFrame) -> list:
    if df is None or df.empty:
        return []
    dd = df.copy()
    dd.columns = [str(c).strip() for c in dd.columns]
    # Prefer Active=Yes if present
    try:
        if "Active" in dd.columns:
            m = dd["Active"].astype(str).str.lower().isin(["yes","true","1"])
            dd_use = dd[m] if m.any() else dd
        else:
            dd_use = dd
    except Exception:
        dd_use = dd
    def pick(colnames: List[str]):
        norm = {str(c).strip().lower(): c for c in dd_use.columns}
        for name in colnames:
            if name in norm and dd_use[norm[name]].notna().any():
                return dd_use[norm[name]].astype(str).str.strip()
        return None
    priority = ["callsign","call sign","unitnumber","unit","unit #","unit_number","name","apparatus","truck"]
    buckets = []
    s = pick(priority)
    if s is not None: buckets.append(s)
    for alt in ["unitnumber","unit","unit #","unit_number","call sign","name","apparatus","truck"]:
        ss = pick([alt])
        if ss is not None: buckets.append(ss)
    if not buckets:
        return []
    s_all = pd.concat(buckets, ignore_index=True)
    vals = (s_all.dropna().map(lambda x: x.strip()).replace("", pd.NA).dropna().unique().tolist())
    return sorted(set(vals))

def repair_rosters(data: Dict[str, pd.DataFrame]) -> Dict[str, pd.DataFrame]:
    p = ensure_columns(data.get("Personnel", pd.DataFrame()), PERSONNEL_SCHEMA).copy()
    if "Rank" in p.columns and not p.empty:
        p["Rank"] = p["Rank"].astype(str)  # free-text ranks allowed
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

# ---------- Auth / permissions ----------
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

# ---------- Sign-in UI ----------
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

# Base tables
ensure_table(data, "Incidents", [
    PRIMARY_KEY,"IncidentDate","IncidentTime","IncidentType","ResponsePriority","AlarmLevel","Shift",
    "LocationName","Address","City","State","PostalCode","Latitude","Longitude",
    "Narrative","Status","CreatedBy","SubmittedAt","ReviewedBy","ReviewedAt","ReviewerComments",
    # extras for print (safe if unused)
    "CallerName","CallerPhone","ReportWriter","Approver"
])
ensure_table(data, "Personnel", PERSONNEL_SCHEMA)
ensure_table(data, "Apparatus", APPARATUS_SCHEMA)
for t, cols in CHILD_TABLES.items():
    ensure_table(data, t, cols)

# Users sheet with defaults
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

# Prepare data & lookups
data = repair_rosters(data)
lookups = get_lookups(data)

if "user" not in st.session_state:
    sign_in_ui(data["Users"]); st.stop()
user = st.session_state["user"]
st.sidebar.write(f"**Logged in as:** {user.get('FullName', user.get('Username',''))}  \\nRole: {user.get('Role','')}")
sign_out_button()

# ---------- Print/Review renderers + PDF export (inline) ----------
def render_incident_block(data: Dict[str, pd.DataFrame], sel):
    rec_df = data.get("Incidents", pd.DataFrame())
    if rec_df.empty:
        st.warning("Incidents table empty."); return
    rec = rec_df[rec_df[PRIMARY_KEY].astype(str) == str(sel)]
    if rec.empty:
        st.warning("Incident not found."); return
    rec = rec.iloc[0].to_dict()

    # Times
    times_df = ensure_columns(data.get("Incident_Times", pd.DataFrame()), ["IncidentNumber","Alarm","Enroute","Arrival","Clear"])
    times_row = {}
    if not times_df.empty:
        m = times_df[PRIMARY_KEY].astype(str) == str(sel)
        if m.any():
            times_row = times_df[m].iloc[0].to_dict()

    st.subheader(f"Incident {sel}")
    st.write(
        f"**Type:** {rec.get('IncidentType','')}  |  "
        f"**Priority:** {rec.get('ResponsePriority','')}  |  "
        f"**Alarm Level:** {rec.get('AlarmLevel','')}"
    )
    st.write(f"**Date:** {rec.get('IncidentDate','')}  **Time:** {rec.get('IncidentTime','')}")
    st.write(
        f"**Location:** {rec.get('LocationName','')} — "
        f"{rec.get('Address','')} {rec.get('City','')} {rec.get('State','')} {rec.get('PostalCode','')}"
    )

    caller_name  = rec.get('CallerName','')
    caller_phone = rec.get('CallerPhone','')
    writer_name  = rec.get('ReportWriter','') or rec.get('CreatedBy','')
    approver     = rec.get('Approver','') or rec.get('ReviewedBy','')
    st.write(
        f"**Caller:** {caller_name if caller_name else '_N/A_'}  |  "
        f"**Caller Phone:** {caller_phone if caller_phone else '_N/A_'}"
    )
    st.write(
        f"**Report Written By:** {writer_name if writer_name else '_N/A_'}  |  "
        f"**Approved By:** {approver if approver else '_N/A_'}"
        f"{' — at ' + str(rec.get('ReviewedAt')) if rec.get('ReviewedAt') else ''}"
    )

    st.write(
        f"**Times —** "
        f"Alarm: {times_row.get('Alarm','')}  |  "
        f"Enroute: {times_row.get('Enroute','')}  |  "
        f"Arrival: {times_row.get('Arrival','')}  |  "
        f"Clear: {times_row.get('Clear','')}"
    )

    st.write("**Narrative:**")
    st.text_area("Narrative (read-only)",
                 value=str(rec.get("Narrative","")),
                 height=220,
                 key=f"narrative_readonly_{sel}",
                 disabled=True)

    ip = ensure_columns(data.get("Incident_Personnel", pd.DataFrame()), ["IncidentNumber","Name","Role","Hours","RespondedIn"])
    ia = ensure_columns(data.get("Incident_Apparatus", pd.DataFrame()), ["IncidentNumber","Unit","UnitType","Role","Actions"])
    ip_view = ip[ip[PRIMARY_KEY].astype(str) == str(sel)]
    ia_view = ia[ia[PRIMARY_KEY].astype(str) == str(sel)]

    st.markdown(f"**Personnel on Scene ({len(ip_view)}):**")
    person_cols = [c for c in ["Name","Role","Hours","RespondedIn"] if c in ip_view.columns]
    st.dataframe(ip_view[person_cols] if not ip_view.empty else ip_view, use_container_width=True, hide_index=True, key=f"grid_print_personnel_{sel}")

    st.markdown(f"**Apparatus on Scene ({len(ia_view)}):**")
    app_cols = [c for c in ["Unit","UnitType","Role","Actions"] if c in ia_view.columns]
    st.dataframe(ia_view[app_cols] if not ia_view.empty else ia_view, use_container_width=True, hide_index=True, key=f"grid_print_apparatus_{sel}")

    # minimal print stylesheet so Streamlit chrome hides during print
    components.html(\"\"\"
    <style>
    @media print {
      header, footer, [data-testid="stSidebar"], .stButton, .stTextInput, .stSlider, .stSelectbox { display: none !important; }
      .block-container { padding: 0 !important; }
    }
    </style>
    \"\"\", height=0, width=0)

def render_print_button(label: str = "🖨️ Print Report"):
    if st.button(label, key="btn_print_report_unified"):
        components.html("<script>window.print()</script>", height=0, width=0)

# ----- PDF export (reportlab optional; HTML fallback) -----
try:
    from reportlab.lib.pagesizes import LETTER
    from reportlab.pdfgen import canvas
    from reportlab.lib.units import inch
    REPORTLAB_OK = True
except Exception:
    REPORTLAB_OK = False

def _wrap_lines(text: str, max_chars: int) -> list:
    if not text:
        return []
    from textwrap import wrap
    return wrap(str(text), max_chars)

def _draw_wrapped(c, text: str, x: float, y: float, max_chars: int, leading: float):
    for line in _wrap_lines(text, max_chars):
        c.drawString(x, y, line)
        y -= leading
        if y < 72:  # 1 inch bottom margin
            c.showPage()
            y = 720
    return y

def _draw_table(c, headers, rows, x, y, col_widths, leading):
    c.setFont("Helvetica", 10)
    y -= leading
    for i, h in enumerate(headers):
        c.drawString(x + sum(col_widths[:i]), y, str(h))
    y -= leading * 0.5
    c.line(x, y, x + sum(col_widths), y)
    y -= leading
    for row in rows:
        for i, cell in enumerate(row):
            c.drawString(x + sum(col_widths[:i]), y, str(cell) if cell is not None else "")
        y -= leading
        if y < 72:
            c.showPage()
            y = 720
    return y

def _generate_pdf_bytes(incident, ip_view, ia_view, times_row, incident_id):
    buf = io.BytesIO()
    c = canvas.Canvas(buf, pagesize=LETTER)
    width, height = LETTER
    left = 0.75 * inch
    y = height - 0.75 * inch

    c.setFont("Helvetica-Bold", 14)
    c.drawString(left, y, f"Incident Report — {incident_id}")
    y -= 18
    c.setFont("Helvetica", 9)
    c.drawString(left, y, f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    y -= 10

    c.setFont("Helvetica", 10)
    y -= 4
    y = _draw_wrapped(c, f"Type: {incident.get('IncidentType','')}  |  Priority: {incident.get('ResponsePriority','')}  |  Alarm Level: {incident.get('AlarmLevel','')}", left, y, 110, 12)
    y = _draw_wrapped(c, f"Date: {incident.get('IncidentDate','')}    Time: {incident.get('IncidentTime','')}", left, y, 110, 12)
    loc = f"{incident.get('LocationName','')} — {incident.get('Address','')} {incident.get('City','')} {incident.get('State','')} {incident.get('PostalCode','')}"
    y = _draw_wrapped(c, f"Location: {loc}", left, y, 110, 12)
    caller = f"Caller: {incident.get('CallerName','') or 'N/A'}   |   Caller Phone: {incident.get('CallerPhone','') or 'N/A'}"
    y = _draw_wrapped(c, caller, left, y, 110, 12)
    writer = f"Report Written By: {incident.get('ReportWriter','') or incident.get('CreatedBy','') or 'N/A'}"
    approver = f"Approved By: {incident.get('Approver','') or incident.get('ReviewedBy','') or 'N/A'}"
    y = _draw_wrapped(c, f"{writer}   |   {approver}", left, y, 110, 12)
    times_line = f"Times — Alarm: {times_row.get('Alarm','')}  |  Enroute: {times_row.get('Enroute','')}  |  Arrival: {times_row.get('Arrival','')}  |  Clear: {times_row.get('Clear','')}"
    y = _draw_wrapped(c, times_line, left, y, 110, 12)

    y -= 6
    c.setFont("Helvetica-Bold", 12)
    c.drawString(left, y, "Narrative")
    y -= 14
    c.setFont("Helvetica", 10)
    y = _draw_wrapped(c, incident.get("Narrative",""), left, y, 110, 12)

    # Personnel table
    y -= 10
    c.setFont("Helvetica-Bold", 12)
    c.drawString(left, y, "Personnel on Scene")
    y -= 14
    c.setFont("Helvetica", 10)
    person_headers = ["Name","Role","Hours","RespondedIn"]
    person_cols = [1.8*inch, 1.3*inch, 0.8*inch, 1.6*inch]
    rows = []
    if not ip_view.empty:
        for _, r in ip_view.iterrows():
            rows.append([r.get("Name",""), r.get("Role",""), r.get("Hours",""), r.get("RespondedIn","")])
    y = _draw_table(c, person_headers, rows, left, y, person_cols, 12)

    # Apparatus table
    y -= 10
    c.setFont("Helvetica-Bold", 12)
    c.drawString(left, y, "Apparatus on Scene")
    y -= 14
    c.setFont("Helvetica", 10)
    app_headers = ["Unit","UnitType","Role","Actions"]
    app_cols = [1.2*inch, 1.2*inch, 1.2*inch, 2.7*inch]
    arows = []
    if not ia_view.empty:
        for _, r in ia_view.iterrows():
            arows.append([r.get("Unit",""), r.get("UnitType",""), r.get("Role",""), r.get("Actions","")])
    y = _draw_table(c, app_headers, arows, left, y, app_cols, 12)

    c.showPage()
    c.save()
    return buf.getvalue()

def _generate_html_bytes(incident, ip_view, ia_view, times_row, incident_id):
    import html
    def esc(s): return html.escape("" if s is None else str(s))
    rows_person = ""
    if not ip_view.empty:
        for _, r in ip_view.iterrows():
            rows_person += f"<tr><td>{esc(r.get('Name',''))}</td><td>{esc(r.get('Role',''))}</td><td>{esc(r.get('Hours',''))}</td><td>{esc(r.get('RespondedIn',''))}</td></tr>"
    rows_app = ""
    if not ia_view.empty:
        for _, r in ia_view.iterrows():
            rows_app += f"<tr><td>{esc(r.get('Unit',''))}</td><td>{esc(r.get('UnitType',''))}</td><td>{esc(r.get('Role',''))}</td><td>{esc(r.get('Actions',''))}</td></tr>"
    html_doc = f"""<!doctype html>
<html><head>
<meta charset='utf-8'>
<title>Incident {esc(incident_id)}</title>
<style>
  body {{ font-family: Arial, Helvetica, sans-serif; margin: 24px; }}
  h1 {{ font-size: 18px; margin: 0 0 8px 0; }}
  h2 {{ font-size: 16px; margin: 18px 0 6px; }}
  .meta {{ margin: 6px 0; }}
  table {{ border-collapse: collapse; width: 100%; }}
  th, td {{ border: 1px solid #999; padding: 6px; font-size: 12px; vertical-align: top; }}
  .muted {{ color: #666; }}
  @media print {{ .noprint {{ display: none; }} }}
</style>
</head>
<body>
  <div class="noprint" style="text-align:right">
    <button onclick="window.print()">Print</button>
  </div>
  <h1>Incident Report — {esc(incident_id)}</h1>
  <div class="muted">Generated {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}</div>
  <div class="meta"><b>Type:</b> {esc(incident.get('IncidentType',''))} &nbsp;&nbsp; <b>Priority:</b> {esc(incident.get('ResponsePriority',''))} &nbsp;&nbsp; <b>Alarm Level:</b> {esc(incident.get('AlarmLevel',''))}</div>
  <div class="meta"><b>Date:</b> {esc(incident.get('IncidentDate',''))} &nbsp;&nbsp; <b>Time:</b> {esc(incident.get('IncidentTime',''))}</div>
  <div class="meta"><b>Location:</b> {esc(incident.get('LocationName',''))} — {esc(incident.get('Address',''))} {esc(incident.get('City',''))} {esc(incident.get('State',''))} {esc(incident.get('PostalCode',''))}</div>
  <div class="meta"><b>Caller:</b> {esc(incident.get('CallerName','') or 'N/A')} &nbsp;&nbsp; <b>Caller Phone:</b> {esc(incident.get('CallerPhone','') or 'N/A')}</div>
  <div class="meta"><b>Report Written By:</b> {esc(incident.get('ReportWriter','') or incident.get('CreatedBy','') or 'N/A')} &nbsp;&nbsp; <b>Approved By:</b> {esc(incident.get('Approver','') or incident.get('ReviewedBy','') or 'N/A')}</div>
  <div class="meta"><b>Times —</b> Alarm: {esc(times_row.get('Alarm',''))} &nbsp; | &nbsp; Enroute: {esc(times_row.get('Enroute',''))} &nbsp; | &nbsp; Arrival: {esc(times_row.get('Arrival',''))} &nbsp; | &nbsp; Clear: {esc(times_row.get('Clear',''))}</div>

  <h2>Narrative</h2>
  <div style="white-space: pre-wrap; font-size: 13px;">{esc(incident.get('Narrative',''))}</div>

  <h2>Personnel on Scene</h2>
  <table>
    <thead><tr><th>Name</th><th>Role</th><th>Hours</th><th>Responded In</th></tr></thead>
    <tbody>{rows_person}</tbody>
  </table>

  <h2>Apparatus on Scene</h2>
  <table>
    <thead><tr><th>Unit</th><th>Unit Type</th><th>Role</th><th>Actions</th></tr></thead>
    <tbody>{rows_app}</tbody>
  </table>
</body></html>"""
    return html_doc.encode("utf-8")

def render_incident_pdf_ui(data: Dict[str, pd.DataFrame], sel):
    incident_df = data.get("Incidents", pd.DataFrame())
    rec = incident_df[incident_df[PRIMARY_KEY].astype(str) == str(sel)]
    if rec.empty:
        st.warning("Incident not found."); return
    incident = rec.iloc[0].to_dict()

    times_df = ensure_columns(data.get("Incident_Times", pd.DataFrame()), ["IncidentNumber","Alarm","Enroute","Arrival","Clear"])
    times_row = {}
    if not times_df.empty:
        m = times_df[PRIMARY_KEY].astype(str) == str(sel)
        if m.any(): times_row = times_df[m].iloc[0].to_dict()

    ip = ensure_columns(data.get("Incident_Personnel", pd.DataFrame()), ["IncidentNumber","Name","Role","Hours","RespondedIn"])
    ia = ensure_columns(data.get("Incident_Apparatus", pd.DataFrame()), ["IncidentNumber","Unit","UnitType","Role","Actions"])
    ip_view = ip[ip[PRIMARY_KEY].astype(str) == str(sel)]
    ia_view = ia[ia[PRIMARY_KEY].astype(str) == str(sel)]

    st.subheader(f"Incident {sel} — Export")
    col1, col2 = st.columns(2)

    if REPORTLAB_OK and col1.button("📄 Convert to PDF", key=f"btn_pdf_{sel}"):
        try:
            pdf_bytes = _generate_pdf_bytes(incident, ip_view, ia_view, times_row, str(sel))
            st.download_button("Download PDF", data=pdf_bytes, file_name=f"incident_{sel}.pdf", mime="application/pdf", key=f"dl_pdf_{sel}")
        except Exception as e:
            st.error(f"PDF generation failed: {e}")

    if col2.button("⬇️ Download HTML (print to PDF)", key=f"btn_html_{sel}"):
        html_bytes = _generate_html_bytes(incident, ip_view, ia_view, times_row, str(sel))
        st.download_button("Download HTML", data=html_bytes, file_name=f"incident_{sel}.html", mime="text/html", key=f"dl_html_{sel}")

# ---------- Tabs ----------
tabs = st.tabs(["Write Report","Review Queue","Rejected","Approved","Rosters","Print","Export","Admin","Diagnostics"])

with tabs[0]:
    st.header("Write Report")
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
            options_df = master.iloc[0:0]
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
                new = [{
                    PRIMARY_KEY: inc_key,
                    "Name": n,
                    "Role": role_default,
                    "Hours": hours_default,
                    "RespondedIn": (responded_in_default or None)
                } for n in picked_people]
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
                new = [{
                    PRIMARY_KEY: inc_key,
                    "Unit": u,
                    "UnitType": (unit_type if unit_type else None),
                    "Role": unit_role,
                    "Actions": unit_actions or ""
                } for u in picked_units]
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
        render_incident_block(data, sel)
        comments = st.text_area("Reviewer Comments", key="rev_comments_queue_auth")
        c = st.columns(3)
        if can(user,"CanReview"):
            if c[0].button("Approve", key="btn_approve_queue_auth"):
                if not can(user,"CanApprove"):
                    st.error("No permission to approve.")
                else:
                    rec = data["Incidents"][data["Incidents"][PRIMARY_KEY].astype(str) == sel].iloc[0].to_dict()
                    rec["Status"] = "Approved"
                    rec["ReviewedBy"] = user.get("Username")
                    rec["ReviewedAt"] = datetime.now().strftime("%Y-%m-%d %H:%M")
                    rec["ReviewerComments"] = comments
                    data["Incidents"] = upsert_row(data["Incidents"], rec, key=PRIMARY_KEY)
                    if st.session_state.get("autosave", True): save_to_path(data, file_path)
                    st.success("Approved.")
            if c[1].button("Reject", key="btn_reject_queue_auth"):
                rec = data["Incidents"][data["Incidents"][PRIMARY_KEY].astype(str) == sel].iloc[0].to_dict()
                rec["Status"] = "Rejected"
                rec["ReviewedBy"] = user.get("Username")
                rec["ReviewedAt"] = datetime.now().strftime("%Y-%m-%d %H:%M")
                rec["ReviewerComments"] = comments or "Please revise."
                data["Incidents"] = upsert_row(data["Incidents"], rec, key=PRIMARY_KEY)
                if st.session_state.get("autosave", True): save_to_path(data, file_path)
                st.warning("Rejected.")
            if c[2].button("Send back to Draft", key="btn_backtodraft_queue_auth"):
                rec = data["Incidents"][data["Incidents"][PRIMARY_KEY].astype(str) == sel].iloc[0].to_dict()
                rec["Status"] = "Draft"
                rec["ReviewerComments"] = comments
                data["Incidents"] = upsert_row(data["Incidents"], rec, key=PRIMARY_KEY)
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
        c = st.columns(3)
        if c[0].button("Move back to Draft to Edit", key="btn_rejected_to_draft"):
            rec["Status"] = "Draft"
            data["Incidents"] = upsert_row(data["Incidents"], rec, key=PRIMARY_KEY)
            if st.session_state.get("autosave", True): save_to_path(data, file_path)
            st.session_state["edit_incident_preselect"] = str(selr)
            st.session_state["force_edit_mode"] = True
            st.success("Moved to Draft. Go to Write Report → Edit to revise and resubmit.")
        with c[1]:
            render_print_button("🖨️ Print This Report")
        with c[2]:
            render_incident_pdf_ui(data, selr)

with tabs[3]:
    st.header("Approved Reports")
    approved = data["Incidents"][data["Incidents"]["Status"].astype(str) == "Approved"]
    st.dataframe(approved, use_container_width=True, hide_index=True, key="grid_approved_auth")
    sela = None
    if not approved.empty:
        sela = st.selectbox("Pick an Approved Incident", options=approved[PRIMARY_KEY].astype(str).tolist(), index=None, placeholder="Choose...", key="pick_approved_auth")
    if sela:
        render_incident_block(data, sela)
        cc = st.columns(2)
        with cc[0]:
            render_print_button("🖨️ Print This Report")
        with cc[1]:
            render_incident_pdf_ui(data, sela)

with tabs[4]:
    st.header("Rosters")
    st.caption("Edit, then click Save. Rank is free text (letters allowed).")
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

with tabs[5]:
    st.header("Print")
    status = st.selectbox("Filter by Status", options=["","Approved","Submitted","Draft","Rejected"], key="print_status_auth")
    base = data["Incidents"].copy()
    if status: base = base[base["Status"].astype(str) == status]
    st.dataframe(base, use_container_width=True, hide_index=True, key="grid_print_auth")
    selp = None
    if not base.empty:
        selp = st.selectbox("Pick an Incident", options=base[PRIMARY_KEY].astype(str).tolist(), index=None, key="print_pick_auth")
    if selp:
        render_incident_block(data, selp)
        cc = st.columns(2)
        with cc[0]:
            render_print_button("🖨️ Print Report")
        with cc[1]:
            render_incident_pdf_ui(data, selp)

with tabs[6]:
    st.header("Export")
    if st.button("Build Excel for Download", key="btn_build_export_auth"):
        payload = save_workbook_to_bytes(data)
        st.download_button("Download Excel", data=payload, file_name="fire_incident_db_export.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key="download_export_auth")
    if st.button("Overwrite Source File Now", key="btn_overwrite_source_auth"):
        ok, err = save_to_path(data, file_path)
        if ok: st.success(f"*" + file_path + "* updated")
        else: st.error(f"Failed: {err}")

with tabs[7]:
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

with tabs[8]:
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
