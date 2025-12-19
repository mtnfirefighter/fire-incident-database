# incident_add_id_patch.py
# FIX: Populate PersonnelID / ApparatusID when adding to incident rows
# This must be called at ADD-TO-INCIDENT time, not print time.

import pandas as pd

def add_personnel_rows(
    incident_number,
    selected_names,
    personnel_df,
    incident_personnel_df,
    role,
    hours,
    responded_in
):
    rows = []
    for name in selected_names:
        match = personnel_df[personnel_df["Name"] == name]
        pid = None
        if not match.empty:
            pid = match.iloc[0]["PersonnelID"]

        rows.append({
            "IncidentNumber": incident_number,
            "PersonnelID": pid,
            "Name": name,
            "Role": role,
            "Hours": hours,
            "RespondedIn": responded_in,
            "Notes": None,
        })

    return pd.concat([incident_personnel_df, pd.DataFrame(rows)], ignore_index=True)


def add_apparatus_rows(
    incident_number,
    selected_units,
    apparatus_df,
    incident_apparatus_df,
    unit_type,
    role,
    actions
):
    rows = []
    for unit in selected_units:
        match = apparatus_df[apparatus_df["Unit"] == unit]
        aid = None
        if not match.empty:
            aid = match.iloc[0]["ApparatusID"]

        rows.append({
            "IncidentNumber": incident_number,
            "ApparatusID": aid,
            "Unit": unit,
            "UnitType": unit_type,
            "Role": role,
            "Actions": actions,
            "Notes": None,
        })

    return pd.concat([incident_apparatus_df, pd.DataFrame(rows)], ignore_index=True)