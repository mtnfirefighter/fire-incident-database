# diag_print_controls.py
from typing import Dict, List, Optional
import io, datetime as _dt
import pandas as pd
import streamlit as st
import streamlit.components.v1 as components

# Try to import reportlab for PDFs; fall back to HTML if missing
try:
    from reportlab.lib.pagesizes import LETTER
    from reportlab.pdfgen import canvas
    from reportlab.lib.units import inch
    _PDF_OK = True
except Exception:
    _PDF_OK = False

def debug_loaded(st):
    st.success("Diagnostic controls loaded ✔", icon="✅")

def _ensure_columns(df: pd.DataFrame, cols: List[str]):
    if df is None:
        df = pd.DataFrame()
    for c in cols:
        if c not in df.columns:
            df[c] = pd.NA
    return df

def _get_rec(data: Dict[str, pd.DataFrame], pk: str, sel) -> Optional[dict]:
    df = data.get("Incidents", pd.DataFrame())
    if df.empty or sel in (None, "", pd.NA):
        return None
    m = df[pk].astype(str) == str(sel)
    if not m.any():
        return None
    return df[m].iloc[0].to_dict()

def _fetch_times(data: Dict[str, pd.DataFrame], pk: str, sel, ensure_columns):
    times = ensure_columns(data.get("Incident_Times", pd.DataFrame()), ["IncidentNumber","Alarm","Enroute","Arrival","Clear"])
    if times.empty:
        return {}
    m = times[pk].astype(str) == str(sel)
    if not m.any():
        return {}
    return times[m].iloc[0].to_dict()

def _pdf_bytes(incident, ip_view, ia_view, times_row, incident_id):
    buf = io.BytesIO()
    c = canvas.Canvas(buf, pagesize=LETTER)
    width, height = LETTER
    left = 0.75 * inch
    y = height - 0.75 * inch

    def _wrap(text, max_chars=110):
        from textwrap import wrap
        return wrap("" if text is None else str(text), max_chars)

    def _draw(lines, font=("Helvetica",10), leading=12):
        nonlocal y
        c.setFont(*font)
        for line in (lines if isinstance(lines, list) else [lines]):
            c.drawString(left, y, line)
            y -= leading
            if y < 72:
                c.showPage(); y = height - 0.75 * inch

    _draw([f"Incident Report — {incident_id}"], font=("Helvetica-Bold",14), leading=18)
    _draw([f"Generated: {_dt.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}"], font=("Helvetica",9), leading=12)
    _draw([""])

    _draw(_wrap(f"Type: {incident.get('IncidentType','')}  |  Priority: {incident.get('ResponsePriority','')}  |  Alarm Level: {incident.get('AlarmLevel','')}"))
    _draw(_wrap(f"Date: {incident.get('IncidentDate','')}    Time: {incident.get('IncidentTime','')}"))
    loc = f"{incident.get('LocationName','')} — {incident.get('Address','')} {incident.get('City','')} {incident.get('State','')} {incident.get('PostalCode','')}"
    _draw(_wrap(f"Location: {loc}"))
    _draw(_wrap(f"Caller: {incident.get('CallerName','') or 'N/A'}   |   Caller Phone: {incident.get('CallerPhone','') or 'N/A'}"))
    _draw(_wrap(f"Report Written By: {incident.get('ReportWriter','') or incident.get('CreatedBy','') or 'N/A'}   |   Approved By: {incident.get('Approver','') or incident.get('ReviewedBy','') or 'N/A'}"))
    _draw(_wrap(f"Times — Alarm: {times_row.get('Alarm','')}  |  Enroute: {times_row.get('Enroute','')}  |  Arrival: {times_row.get('Arrival','')}  |  Clear: {times_row.get('Clear','')}"))
    _draw([""])
    _draw(["Narrative"], font=("Helvetica-Bold",12), leading=14)
    _draw(_wrap(incident.get("Narrative","")), font=("Helvetica",10), leading=12)

    # Simple tables
    def _table(title, headers, rows):
        _draw([""], leading=10)
        _draw([title], font=("Helvetica-Bold",12), leading=14)
        _draw([" | ".join(headers)], font=("Helvetica",10))
        _draw(["-"*92], font=("Helvetica",10))
        for row in rows:
            _draw([" | ".join("" if v is None else str(v) for v in row)], font=("Helvetica",10))

    # Personnel
    prows = []
    if ip_view is not None and not ip_view.empty:
        for _, r in ip_view.iterrows():
            prows.append([r.get("Name",""), r.get("Role",""), r.get("Hours",""), r.get("RespondedIn","")])
    _table("Personnel on Scene", ["Name","Role","Hours","RespondedIn"], prows)

    # Apparatus
    arows = []
    if ia_view is not None and not ia_view.empty:
        for _, r in ia_view.iterrows():
            arows.append([r.get("Unit",""), r.get("UnitType",""), r.get("Role",""), r.get("Actions","")])
    _table("Apparatus on Scene", ["Unit","UnitType","Role","Actions"], arows)

    c.showPage(); c.save()
    return buf.getvalue()

def _html_bytes(incident, ip_view, ia_view, times_row, incident_id):
    import html
    def esc(x): return html.escape("" if x is None else str(x))
    def rows(df, cols):
        if df is None or df.empty: return ""
        out = []
        for _, r in df.iterrows():
            out.append("<tr>" + "".join(f"<td>{esc(r.get(c,''))}</td>" for c in cols) + "</tr>")
        return "\n".join(out)

    html_doc = f\"\"\"<!doctype html>
<html><head><meta charset="utf-8"><title>Incident {esc(incident_id)}</title>
<style>
 body {{ font-family: Arial, Helvetica, sans-serif; margin: 24px; }}
 table {{ border-collapse: collapse; width: 100%; }}
 th, td {{ border: 1px solid #999; padding: 6px; font-size: 12px; vertical-align: top; }}
 .meta {{ margin: 6px 0; }}
 .muted {{ color: #666; }}
 @media print {{ .noprint {{ display: none; }} }}
</style></head>
<body>
 <div class="noprint" style="text-align:right"><button onclick="window.print()">Print</button></div>
 <h1>Incident Report — {esc(incident_id)}</h1>
 <div class="muted">Generated {_dt.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}</div>
 <div class="meta"><b>Type:</b> {esc(incident.get('IncidentType',''))} &nbsp;&nbsp; <b>Priority:</b> {esc(incident.get('ResponsePriority',''))} &nbsp;&nbsp; <b>Alarm Level:</b> {esc(incident.get('AlarmLevel',''))}</div>
 <div class="meta"><b>Date:</b> {esc(incident.get('IncidentDate',''))} &nbsp;&nbsp; <b>Time:</b> {esc(incident.get('IncidentTime',''))}</div>
 <div class="meta"><b>Location:</b> {esc(incident.get('LocationName',''))} — {esc(incident.get('Address',''))} {esc(incident.get('City',''))} {esc(incident.get('State',''))} {esc(incident.get('PostalCode',''))}</div>
 <div class="meta"><b>Caller:</b> {esc(incident.get('CallerName','') or 'N/A')} &nbsp;&nbsp; <b>Caller Phone:</b> {esc(incident.get('CallerPhone','') or 'N/A')}</div>
 <div class="meta"><b>Report Written By:</b> {esc(incident.get('ReportWriter','') or incident.get('CreatedBy','') or 'N/A')} &nbsp;&nbsp; <b>Approved By:</b> {esc(incident.get('Approver','') or incident.get('ReviewedBy','') or 'N/A')}</div>
 <div class="meta"><b>Times —</b> Alarm: {esc(times_row.get('Alarm',''))} &nbsp; | &nbsp; Enroute: {esc(times_row.get('Enroute',''))} &nbsp; | &nbsp; Arrival: {esc(times_row.get('Arrival',''))} &nbsp; | &nbsp; Clear: {esc(times_row.get('Clear',''))}</div>
 <h2>Narrative</h2>
 <div style="white-space: pre-wrap; font-size: 13px;">{esc(incident.get('Narrative',''))}</div>
 <h2>Personnel on Scene</h2>
 <table><thead><tr><th>Name</th><th>Role</th><th>Hours</th><th>Responded In</th></tr></thead>
 <tbody>{rows(ip_view,['Name','Role','Hours','RespondedIn'])}</tbody></table>
 <h2>Apparatus on Scene</h2>
 <table><thead><tr><th>Unit</th><th>Unit Type</th><th>Role</th><th>Actions</th></tr></thead>
 <tbody>{rows(ia_view,['Unit','UnitType','Role','Actions'])}</tbody></table>
</body></html>\"\"\"
    return html_doc.encode("utf-8")

def print_controls_block(st, data: Dict[str, pd.DataFrame], PRIMARY_KEY: str, selected_id, ensure_columns, area_key: str = "default"):
    \"\"\"Render always-visible Print + PDF controls. Pass your selected id variable.
    area_key: short string that scopes Streamlit keys per tab (e.g., 'print_tab', 'review_tab').
    \"\"\"
    st.info(f"**Selected ID:** {selected_id if selected_id else '— none —'}")

    if not selected_id:
        st.warning("Pick an incident in the dropdown above to enable export/print.")
        return

    rec = _get_rec(data, PRIMARY_KEY, selected_id)
    if rec is None:
        st.error("Selected id not found in Incidents.")
        return

    ip = ensure_columns(data.get("Incident_Personnel", pd.DataFrame()), ["IncidentNumber","Name","Role","Hours","RespondedIn"])
    ia = ensure_columns(data.get("Incident_Apparatus", pd.DataFrame()), ["IncidentNumber","Unit","UnitType","Role","Actions"])
    ip_view = ip[ip[PRIMARY_KEY].astype(str) == str(selected_id)]
    ia_view = ia[ia[PRIMARY_KEY].astype(str) == str(selected_id)]
    times_row = _fetch_times(data, PRIMARY_KEY, selected_id, ensure_columns)

    # Buttons with unique keys per tab + id
    c1, c2, c3 = st.columns(3)
    if c1.button("🖨️ Print Page", key=f"btn_print_{area_key}_{selected_id}"):
        components.html("<script>window.print()</script>", height=0, width=0)

    if _PDF_OK and c2.button("📄 Convert to PDF", key=f"btn_pdf_{area_key}_{selected_id}"):
        try:
            pdf = _pdf_bytes(rec, ip_view, ia_view, times_row, str(selected_id))
            st.download_button("Download PDF", data=pdf, file_name=f"incident_{selected_id}.pdf", mime="application/pdf",
                               key=f"dl_pdf_{area_key}_{selected_id}")
        except Exception as e:
            st.error(f"PDF failed: {e}")

    if c3.button("⬇️ Download HTML", key=f"btn_html_{area_key}_{selected_id}"):
        html = _html_bytes(rec, ip_view, ia_view, times_row, str(selected_id))
        st.download_button("Download HTML", data=html, file_name=f"incident_{selected_id}.html", mime="text/html",
                           key=f"dl_html_{area_key}_{selected_id}")
