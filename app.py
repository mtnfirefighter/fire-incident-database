import os, io
from datetime import datetime, date
from typing import Dict, List

import pandas as pd
import streamlit as st

st.set_page_config(page_title='Fire Incident Reports', page_icon='📝', layout='wide')

DEFAULT_FILE = os.path.join(os.path.dirname(__file__), 'fire_incident_db.xlsx')
PRIMARY_KEY = 'IncidentNumber'
CHILD_TABLES = {
    'Incident_Times': ['IncidentNumber','Alarm','Enroute','Arrival','Clear'],
    'Incident_Personnel': ['IncidentNumber','Name','Role','Hours'],
    'Incident_Apparatus': ['IncidentNumber','Unit','UnitType','Role','Actions'],
    'Incident_Actions': ['IncidentNumber','Action','Notes'],
}
PERSONNEL_SCHEMA = ['PersonnelID','Name','UnitNumber','Rank','Badge','Phone','Email','Address','City','State','PostalCode','Certifications','Active','FirstName','LastName','FullName']
APPARATUS_SCHEMA = ['ApparatusID','UnitNumber','CallSign','UnitType','GPM','TankSize','SeatingCapacity','Station','Active','Name']
USERS_SCHEMA = ['Username','Password','Role','FullName','Active']

LOOKUP_SHEETS = {
    'List_IncidentType': 'IncidentType',
    'List_AlarmLevel': 'AlarmLevel',
    'List_ResponsePriority': 'ResponsePriority',
    'List_PersonnelRoles': 'Role',
    'List_UnitTypes': 'UnitType',
    'List_Actions': 'Action',
    'List_States': 'State',
}


def load_workbook(path: str) -> Dict[str, pd.DataFrame]:
    try:
        xls = pd.ExcelFile(path)
        return {name: xls.parse(name) for name in xls.sheet_names}
    except Exception as e:
        st.error(f'Failed to load: {e}')
        return {}

def save_workbook_to_bytes(dfs: Dict[str, pd.DataFrame]) -> bytes:
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine='xlsxwriter') as writer:
        for sheet, df in dfs.items():
            df.to_excel(writer, sheet_name=sheet, index=False)
    buf.seek(0)
    return buf.read()

def save_to_path(dfs: Dict[str, pd.DataFrame], path: str):
    try:
        with pd.ExcelWriter(path, engine='xlsxwriter') as writer:
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


def build_names(df: pd.DataFrame) -> list:
    if 'Name' in df and df['Name'].notna().any():
        s = df['Name'].astype(str)
    elif 'FullName' in df and df['FullName'].notna().any():
        s = df['FullName'].astype(str)
    elif all(c in df.columns for c in ['FirstName','LastName','Rank']):
        s = (df['Rank'].fillna('').astype(str).str.strip() + ' ' + df['FirstName'].fillna('').astype(str).str.strip() + ' ' + df['LastName'].fillna('').astype(str).str.strip()).str.replace(r'\s+',' ', regex=True).str.strip()
    elif all(c in df.columns for c in ['FirstName','LastName']):
        s = (df['FirstName'].fillna('').astype(str).str.strip() + ' ' + df['LastName'].fillna('').astype(str).str.strip()).str.strip()
    else:
        s = pd.Series([], dtype=str)
    vals = s.dropna().map(lambda x: x.strip()).replace('', pd.NA).dropna().unique().tolist()
    return sorted(set(vals))

def build_units(df: pd.DataFrame) -> list:
    for col in ['UnitNumber','CallSign','Name']:
        if col in df.columns and df[col].notna().any():
            s = df[col].astype(str); break
    else:
        s = pd.Series([], dtype=str)
    vals = s.dropna().map(lambda x: x.strip()).replace('', pd.NA).dropna().unique().tolist()
    return sorted(set(vals))

# ---------------- App Boot ----------------
st.sidebar.title('📝 Fire Incident Reports (v4.2.6)')
file_path = st.sidebar.text_input('Excel path', value=DEFAULT_FILE, key='sidebar_path')
uploaded = st.sidebar.file_uploader('Upload/replace workbook (.xlsx)', type=['xlsx'], key='sidebar_upload')
if uploaded:
    with open(file_path, 'wb') as f: f.write(uploaded.read())
    st.sidebar.success(f'Saved to {file_path}')
st.session_state.setdefault('autosave', True)
st.session_state['autosave'] = st.sidebar.toggle('Autosave to Excel', value=True)
st.session_state.setdefault('roster_source', 'app')
st.session_state['roster_source'] = st.sidebar.selectbox('Roster source', options=['app','excel'], index=0, help='Use the live in‑app roster (session) or read directly from the Excel file.')

data: Dict[str, pd.DataFrame] = {}
if os.path.exists(file_path): data = load_workbook(file_path)
if not data: st.info('Upload or point to your Excel workbook to begin.'); st.stop()

# Ensure tables exist
if 'Incidents' not in data:
    data['Incidents'] = pd.DataFrame(columns=[
        PRIMARY_KEY,'IncidentDate','IncidentTime','IncidentType','ResponsePriority','AlarmLevel','Shift',
        'LocationName','Address','City','State','PostalCode','Latitude','Longitude',
        'Narrative','Status','CreatedBy','SubmittedAt','ReviewedBy','ReviewedAt','ReviewerComments'
    ])
for t, cols in CHILD_TABLES.items():
    if t not in data: data[t] = pd.DataFrame(columns=cols)
data['Personnel'] = ensure_columns(data.get('Personnel', pd.DataFrame()), PERSONNEL_SCHEMA)
data['Apparatus'] = ensure_columns(data.get('Apparatus', pd.DataFrame()), APPARATUS_SCHEMA)

# Seed session_state rosters if not present
st.session_state.setdefault('roster_personnel', data['Personnel'].copy())
st.session_state.setdefault('roster_apparatus', data['Apparatus'].copy())

lookups = get_lookups(data)

def get_roster(table: str) -> pd.DataFrame:
    use_app = st.session_state.get('roster_source','app') == 'app'
    if use_app:
        return st.session_state['roster_personnel'].copy() if table == 'Personnel' else st.session_state['roster_apparatus'].copy()
    else:
        return data[table].copy()

tabs = st.tabs(['Write Report','Review Queue','Rosters','Print','Export'])

# ---- Write Report
with tabs[0]:
    st.header('Write Report')
    master = data['Incidents']
    mode = st.radio('Mode', ['New','Edit'], horizontal=True, key='mode_write_426')
    defaults = {}; selected = None
    if mode == 'Edit':
        options_df = master
        options = options_df[PRIMARY_KEY].dropna().astype(str).tolist() if PRIMARY_KEY in options_df.columns else []
        selected = st.selectbox('Select IncidentNumber', options=options, index=None, placeholder='Choose...', key='pick_edit_write_426')
        if selected:
            defaults = master[master[PRIMARY_KEY].astype(str) == selected].iloc[0].to_dict()

    with st.container(border=True):
        st.subheader('Incident Details')
        c1, c2, c3 = st.columns(3)
        inc_num = c1.text_input('IncidentNumber', value=str(defaults.get(PRIMARY_KEY,'')) if defaults else '', key='w_inc_num_426')
        inc_date = c2.date_input('IncidentDate', value=pd.to_datetime(defaults.get('IncidentDate')).date() if defaults.get('IncidentDate') is not None and str(defaults.get('IncidentDate')) != 'NaT' else date.today(), key='w_inc_date_426')
        inc_time = c3.text_input('IncidentTime (HH:MM)', value=str(defaults.get('IncidentTime','')) if defaults else '', key='w_inc_time_426')
        c4, c5, c6 = st.columns(3)
        inc_type = c4.selectbox('IncidentType', options=['']+lookups.get('IncidentType', []), index=0, key='w_type_426')
        inc_prio = c5.selectbox('ResponsePriority', options=['']+lookups.get('ResponsePriority', []), index=0, key='w_prio_426')
        inc_alarm = c6.selectbox('AlarmLevel', options=['']+lookups.get('AlarmLevel', []), index=0, key='w_alarm_426')
        c7, c8, c9 = st.columns(3)
        loc_name = c7.text_input('LocationName', value=str(defaults.get('LocationName','')) if defaults else '', key='w_locname_426')
        addr = c8.text_input('Address', value=str(defaults.get('Address','')) if defaults else '', key='w_addr_426')
        city = c9.text_input('City', value=str(defaults.get('City','')) if defaults else '', key='w_city_426')
        c10, c11, c12 = st.columns(3)
        state = c10.text_input('State', value=str(defaults.get('State','')) if defaults else '', key='w_state_426')
        postal = c11.text_input('PostalCode', value=str(defaults.get('PostalCode','')) if defaults else '', key='w_postal_426')
        shift = c12.text_input('Shift', value=str(defaults.get('Shift','')) if defaults else '', key='w_shift_426')

    with st.container(border=True):
        st.subheader('Narrative')
        narrative = st.text_area('Write full narrative here', value=str(defaults.get('Narrative','')) if defaults else '', height=300, key='w_narrative_426')

    with st.container(border=True):
        st.subheader('All Members on Scene')
        roster_people = get_roster('Personnel')
        if 'Name' in roster_people and roster_people['Name'].notna().any():
            names = roster_people['Name'].astype(str)
        elif 'FullName' in roster_people and roster_people['FullName'].notna().any():
            names = roster_people['FullName'].astype(str)
        elif all(c in roster_people.columns for c in ['FirstName','LastName','Rank']):
            names = (roster_people['Rank'].fillna('').astype(str).str.strip() + ' ' + roster_people['FirstName'].fillna('').astype(str).str.strip() + ' ' + roster_people['LastName'].fillna('').astype(str).str.strip()).str.replace(r'\s+',' ', regex=True).str.strip()
        else:
            names = pd.Series([], dtype=str)
        name_opts = sorted(set(names.dropna().map(lambda s: s.strip()).replace('', pd.NA).dropna().unique().tolist()))
        picked_people = st.multiselect('Pick members', options=name_opts, key='w_pick_people_426')
        roles = get_lookups(data).get('Role', ['OIC','Driver','Firefighter'])
        c = st.columns(3)
        role_default = c[0].selectbox('Default Role', options=roles, index=0 if roles else None, key='w_role_default_426')
        hours_default = c[1].number_input('Default Hours', value=0.0, min_value=0.0, step=0.5, key='w_hours_default_426')
        add_people = c[2].button('Add Selected Members', key='w_add_people_btn_426')
        st.caption(f"Current context → IncidentNumber: **{inc_num or '(empty)'}**, Selected members: {len(picked_people)}")
        if add_people:
            if not inc_num or str(inc_num).strip() == '':
                st.error('Enter **IncidentNumber** before adding members.'); st.stop()
            inc_key = str(inc_num).strip()
            df = ensure_columns(data.get('Incident_Personnel', pd.DataFrame()), CHILD_TABLES['Incident_Personnel'])
            new = [{PRIMARY_KEY: inc_key, 'Name': n, 'Role': role_default, 'Hours': hours_default} for n in picked_people]
            if new:
                data['Incident_Personnel'] = pd.concat([df, pd.DataFrame(new)], ignore_index=True)
                if st.session_state.get('autosave', True):
                    ok, err = save_to_path(data, file_path)
                    if not ok: st.error(f'Autosave failed: {err}')
                st.success(f'Added {len(new)} member(s) to incident {inc_key}.')
            else:
                st.warning('No members selected.')
        cur_per = ensure_columns(data.get('Incident_Personnel', pd.DataFrame()), CHILD_TABLES['Incident_Personnel'])
        cur_view = cur_per[cur_per[PRIMARY_KEY].astype(str) == (str(inc_num).strip() if inc_num else '___no_key___')]
        st.write(f"**Total Personnel on Scene:** {0 if cur_view is None or cur_view.empty else len(cur_view)}")
        st.dataframe(cur_view if cur_view is not None else pd.DataFrame(columns=CHILD_TABLES['Incident_Personnel']), use_container_width=True, hide_index=True)

    with st.container(border=True):
        st.subheader('Apparatus on Scene')
        roster_units = get_roster('Apparatus')
        label = None
        for col in ['UnitNumber','CallSign','Name']:
            if col in roster_units.columns and roster_units[col].notna().any(): label = col; break
        unit_series = roster_units[label].astype(str) if label else pd.Series([], dtype=str)
        unit_opts = sorted(set(unit_series.dropna().map(lambda s: s.strip()).replace('', pd.NA).dropna().unique().tolist()))
        picked_units = st.multiselect('Pick apparatus units', options=unit_opts, key='w_pick_units_426')
        c = st.columns(3)
        unit_type = c[0].selectbox('UnitType', options=['']+get_lookups(data).get('UnitType', []), index=0, key='w_unit_type_426')
        unit_role = c[1].selectbox('Role', options=['Primary','Support','Water Supply','Staging'], index=0, key='w_unit_role_426')
        add_units = c[2].button('Add Selected Units', key='w_add_units_btn_426')
        st.caption(f"Current context → IncidentNumber: **{inc_num or '(empty)'}**, Selected units: {len(picked_units)}")
        if add_units:
            if not inc_num or str(inc_num).strip() == '':
                st.error('Enter **IncidentNumber** before adding apparatus.'); st.stop()
            inc_key = str(inc_num).strip()
            df = ensure_columns(data.get('Incident_Apparatus', pd.DataFrame()), CHILD_TABLES['Incident_Apparatus'])
            new = [{PRIMARY_KEY: inc_key, 'Unit': u, 'UnitType': (unit_type if unit_type else None), 'Role': unit_role, 'Actions': ''} for u in picked_units]
            if new:
                data['Incident_Apparatus'] = pd.concat([df, pd.DataFrame(new)], ignore_index=True)
                if st.session_state.get('autosave', True):
                    ok, err = save_to_path(data, file_path)
                    if not ok: st.error(f'Autosave failed: {err}')
                st.success(f'Added {len(new)} unit(s) to incident {inc_key}.')
            else:
                st.warning('No units selected.')
        cur_app = ensure_columns(data.get('Incident_Apparatus', pd.DataFrame()), CHILD_TABLES['Incident_Apparatus'])
        cur_app_view = cur_app[cur_app[PRIMARY_KEY].astype(str) == (str(inc_num).strip() if inc_num else '___no_key___')]
        st.write(f"**Total Apparatus on Scene:** {0 if cur_app_view is None or cur_app_view.empty else len(cur_app_view)}")
        st.dataframe(cur_app_view if cur_app_view is not None else pd.DataFrame(columns=CHILD_TABLES['Incident_Apparatus']), use_container_width=True, hide_index=True)

    with st.container(border=True):
        st.subheader('Times (optional)')
        c = st.columns(4)
        alarm = c[0].text_input('Alarm (HH:MM)', key='w_alarm_time_426')
        enroute = c[1].text_input('Enroute (HH:MM)', key='w_enroute_time_426')
        arrival = c[2].text_input('Arrival (HH:MM)', key='w_arrival_time_426')
        clear = c[3].text_input('Clear (HH:MM)', key='w_clear_time_426')
        if st.button('Save Times', key='w_save_times_426'):
            if not inc_num or str(inc_num).strip() == '':
                st.error('Enter **IncidentNumber** before saving times.')
            else:
                inc_key = str(inc_num).strip()
                times = data['Incident_Times']
                new = {PRIMARY_KEY: inc_key, 'Alarm': alarm, 'Enroute': enroute, 'Arrival': arrival, 'Clear': clear}
                data['Incident_Times'] = upsert_row(times, new, key=PRIMARY_KEY)
                if st.session_state.get('autosave', True):
                    ok, err = save_to_path(data, file_path)
                    if not ok: st.error(f'Autosave failed: {err}')
                st.success('Times saved.')

    row_vals = {
        PRIMARY_KEY: (str(inc_num).strip() if inc_num else ''),
        'IncidentDate': pd.to_datetime(inc_date),
        'IncidentTime': inc_time,
        'IncidentType': inc_type,
        'ResponsePriority': inc_prio,
        'AlarmLevel': inc_alarm,
        'LocationName': loc_name,
        'Address': addr,
        'City': city,
        'State': state,
        'PostalCode': postal,
        'Shift': shift,
        'Narrative': narrative,
        'CreatedBy': 'member',
    }
    actions = st.columns(3)
    if actions[0].button('Save Draft', key='w_save_draft_426'):
        if not row_vals[PRIMARY_KEY]:
            st.error('Enter **IncidentNumber** before saving.'); 
        else:
            row_vals['Status'] = 'Draft'
            data['Incidents'] = upsert_row(data['Incidents'], row_vals, key=PRIMARY_KEY)
            if st.session_state.get('autosave', True):
                ok, err = save_to_path(data, file_path)
                if not ok: st.error(f'Autosave failed: {err}')
            st.success('Draft saved.')
    if actions[1].button('Submit for Review', key='w_submit_review_426'):
        if not row_vals[PRIMARY_KEY]:
            st.error('Enter **IncidentNumber** before submitting.'); 
        else:
            row_vals['Status'] = 'Submitted'; row_vals['SubmittedAt'] = datetime.now().strftime('%Y-%m-%d %H:%M')
            data['Incidents'] = upsert_row(data['Incidents'], row_vals, key=PRIMARY_KEY)
            if st.session_state.get('autosave', True):
                ok, err = save_to_path(data, file_path)
                if not ok: st.error(f'Autosave failed: {err}')
            st.success('Submitted for review.')

# ---- Review Queue
with tabs[1]:
    st.header('Review Queue')
    pending = data['Incidents'][data['Incidents']['Status'].astype(str) == 'Submitted']
    st.dataframe(pending, use_container_width=True, hide_index=True)
    sel = None
    if not pending.empty:
        sel = st.selectbox('Pick an Incident to review', options=pending[PRIMARY_KEY].astype(str).tolist(), index=None, placeholder='Choose...', key='pick_review_queue_426')
    if sel:
        rec = data['Incidents'][data['Incidents'][PRIMARY_KEY].astype(str) == sel].iloc[0].to_dict()
        st.json(rec)
        comments = st.text_area('Reviewer Comments', key='rev_comments_queue_426')
        c = st.columns(3)
        if c[0].button('Approve', key='btn_approve_queue_426'):
            row = rec; row['Status'] = 'Approved'; row['ReviewedBy'] = 'reviewer'; row['ReviewedAt'] = datetime.now().strftime('%Y-%m-%d %H:%M'); row['ReviewerComments'] = comments
            data['Incidents'] = upsert_row(data['Incidents'], row, key=PRIMARY_KEY); 
            if st.session_state.get('autosave', True):
                save_to_path(data, file_path)
            st.success('Approved.')
        if c[1].button('Reject', key='btn_reject_queue_426'):
            row = rec; row['Status'] = 'Rejected'; row['ReviewedBy'] = 'reviewer'; row['ReviewedAt'] = datetime.now().strftime('%Y-%m-%d %H:%M'); row['ReviewerComments'] = comments or 'Please revise.'
            data['Incidents'] = upsert_row(data['Incidents'], row, key=PRIMARY_KEY); 
            if st.session_state.get('autosave', True):
                save_to_path(data, file_path)
            st.warning('Rejected.')
        if c[2].button('Send back to Draft', key='btn_backtodraft_queue_426'):
            row = rec; row['Status'] = 'Draft'; row['ReviewerComments'] = comments
            data['Incidents'] = upsert_row(data['Incidents'], row, key=PRIMARY_KEY); 
            if st.session_state.get('autosave', True):
                save_to_path(data, file_path)
            st.info('Moved back to Draft.')

# ---- Rosters
with tabs[2]:
    st.header('Rosters')
    st.caption("Edits update the in‑app roster immediately; click Save to write to Excel. Pickers use this in‑app roster when 'Roster source' = app.")
    personnel = ensure_columns(st.session_state['roster_personnel'], PERSONNEL_SCHEMA)
    st.session_state['roster_personnel'] = st.data_editor(personnel, num_rows='dynamic', use_container_width=True, key='editor_personnel_roster_426')
    c = st.columns(2)
    if c[0].button('Save Personnel Roster to Excel', key='save_personnel_roster_426'):
        data['Personnel'] = ensure_columns(st.session_state['roster_personnel'], PERSONNEL_SCHEMA); 
        save_to_path(data, file_path); st.success('Personnel roster saved to Excel.')
    apparatus = ensure_columns(st.session_state['roster_apparatus'], APPARATUS_SCHEMA)
    st.session_state['roster_apparatus'] = st.data_editor(apparatus, num_rows='dynamic', use_container_width=True, key='editor_apparatus_roster_426')
    if c[1].button('Save Apparatus Roster to Excel', key='save_apparatus_roster_426'):
        data['Apparatus'] = ensure_columns(st.session_state['roster_apparatus'], APPARATUS_SCHEMA); 
        save_to_path(data, file_path); st.success('Apparatus roster saved to Excel.')

# ---- Print
with tabs[3]:
    st.header('Print')
    status = st.selectbox('Filter by Status', options=['','Approved','Submitted','Draft','Rejected'], key='print_status_center_426')
    base = data['Incidents'].copy()
    if status: base = base[base['Status'].astype(str) == status]
    st.dataframe(base, use_container_width=True, hide_index=True)
    sel = None
    if not base.empty:
        sel = st.selectbox('Pick an Incident', options=base[PRIMARY_KEY].astype(str).tolist(), index=None, key='print_pick_center_426')
    if sel:
        rec = data['Incidents'][data['Incidents'][PRIMARY_KEY].astype(str) == sel].iloc[0].to_dict()
        st.json(rec)

# ---- Export
with tabs[4]:
    st.header('Export')
    if st.button('Build Excel for Download', key='btn_build_export_426'):
        payload = save_workbook_to_bytes(data)
        st.download_button('Download Excel', data=payload, file_name='fire_incident_db_export.xlsx', mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', key='download_export_426')
    if st.button('Overwrite Source File Now', key='btn_overwrite_source_426'):
        ok, err = save_to_path(data, file_path)
        if ok: st.success(f'Overwrote: {file_path}')
        else: st.error(f'Failed to write: {err}')
