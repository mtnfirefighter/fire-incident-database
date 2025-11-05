# pdf_export_v4_3_2.py
from typing import Dict, List, Optional
import io, datetime
import pandas as pd
from textwrap import wrap
import streamlit as st
try:
    from reportlab.lib.pagesizes import LETTER
    from reportlab.pdfgen import canvas
    from reportlab.lib.units import inch
    REPORTLAB_OK = True
except Exception:
    REPORTLAB_OK = False

def _ensure_columns(df: pd.DataFrame, cols: List[str]):
    if df is None:
        df = pd.DataFrame()
    for c in cols:
        if c not in df.columns:
            df[c] = pd.NA
    return df

def _get_incident_record(data: Dict[str, pd.DataFrame], pk: str, sel) -> Optional[dict]:
    rec_df = data.get("Incidents", pd.DataFrame())
    if rec_df.empty:
        return None
    rec = rec_df[rec_df[pk].astype(str) == str(sel)]
    if rec.empty:
        return None
    return rec.iloc[0].to_dict()

def _fetch_times_row(data: Dict[str, pd.DataFrame], pk: str, sel, ensure_columns):
    times_df = ensure_columns(data.get("Incident_Times", pd.DataFrame()), ["IncidentNumber","Alarm","Enroute","Arrival","Clear"])
    if not times_df.empty:
        match = times_df[times_df[pk].astype(str) == str(sel)]
        if not match.empty:
            return match.iloc[0].to_dict()
    return {}

def _wrap_lines(c, text: str, max_chars: int) -> list:
    if not text:
        return []
    return wrap(str(text), max_chars)

def _draw_wrapped(c, text: str, x: float, y: float, max_chars: int, leading: float):
    for line in _wrap_lines(c, text, max_chars):
        c.drawString(x, y, line)
        y -= leading
        if y < 72:  # 1 inch bottom margin
            c.showPage()
            y = 720
    return y

def _draw_table(c, headers, rows, x, y, col_widths, leading):
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
    from reportlab.lib.pagesizes import LETTER
    from reportlab.pdfgen import canvas
    from reportlab.lib.units import inch

    buf = io.BytesIO()
    c = canvas.Canvas(buf, pagesize=LETTER)
    width, height = LETTER

    left = 0.75 * inch
    y = height - 0.75 * inch

    c.setFont("Helvetica-Bold", 14)
    c.drawString(left, y, f"Incident Report — {incident_id}")
    y -= 18
    c.setFont("Helvetica", 9)
    c.drawString(left, y, f"Generated: {datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    y -= 10

    c.setFont("Helvetica", 10)
    y -= 4
    y = _draw_wrapped(c, f"Type: {incident.get('IncidentType','')}  |  Priority: {incident.get('ResponsePriority','')}  |  Alarm Level: {incident.get('AlarmLevel','')}", left, y, 110, 12)
    y = _draw_wrapped(c, f"Date: {incident.get('IncidentDate','')}    Time: {incident.get('IncidentTime','')}", left, y, 110, 12)
    loc = f\"{incident.get('LocationName','')} — {incident.get('Address','')} {incident.get('City','')} {incident.get('State','')} {incident.get('PostalCode','')}\"
    y = _draw_wrapped(c, f\"Location: {loc}\", left, y, 110, 12)

    caller = f\"Caller: {incident.get('CallerName','') or 'N/A'}   |   Caller Phone: {incident.get('CallerPhone','') or 'N/A'}\"
    y = _draw_wrapped(c, caller, left, y, 110, 12)

    writer = f\"Report Written By: {incident.get('ReportWriter','') or incident.get('CreatedBy','') or 'N/A'}\"
    approver = f\"Approved By: {incident.get('Approver','') or incident.get('ReviewedBy','') or 'N/A'}\"
    y = _draw_wrapped(c, f\"{writer}   |   {approver}\", left, y, 110, 12)

    times_line = f\"Times — Alarm: {times_row.get('Alarm','')}  |  Enroute: {times_row.get('Enroute','')}  |  Arrival: {times_row.get('Arrival','')}  |  Clear: {times_row.get('Clear','')}\"
    y = _draw_wrapped(c, times_line, left, y, 110, 12)

    y -= 6
    c.setFont("Helvetica-Bold", 12)
    c.drawString(left, y, "Narrative")
    y -= 14
    c.setFont("Helvetica", 10)
    y = _draw_wrapped(c, incident.get("Narrative",""), left, y, 110, 12)

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
            rows_person += f\"<tr><td>{esc(r.get('Name',''))}</td><td>{esc(r.get('Role',''))}</td><td>{esc(r.get('Hours',''))}</td><td>{esc(r.get('RespondedIn',''))}</td></tr>\"

    rows_app = ""
    if not ia_view.empty:
        for _, r in ia_view.iterrows():
            rows_app += f\"<tr><td>{esc(r.get('Unit',''))}</td><td>{esc(r.get('UnitType',''))}</td><td>{esc(r.get('Role',''))}</td><td>{esc(r.get('Actions',''))}</td></tr>\"

    html_doc = f\"\"\"<!doctype html>
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
  <div class="muted">Generated {datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}</div>
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
</body></html>\"\"\"
    return html_doc.encode("utf-8")

def render_incident_pdf_ui(st, data: Dict[str, pd.DataFrame], PRIMARY_KEY: str, sel, ensure_columns):
    incident = _get_incident_record(data, PRIMARY_KEY, sel)
    if incident is None:
        st.warning("Incident not found."); return

    times_row = _fetch_times_row(data, PRIMARY_KEY, sel, ensure_columns)
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
