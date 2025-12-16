
# --- PATCH: PersonnelID auto-fill fix ---
# This patch ensures that when adding members to an incident,
# the PersonnelID field is automatically filled from the Personnel roster.

import re
import pandas as pd
import streamlit as st

def add_members_with_id_fix(picked_people, inc_num, role_default, hours_default, responded_in_default, data, PRIMARY_KEY, CHILD_TABLES, PERSONNEL_SCHEMA):
    people_df = data.get("Personnel", pd.DataFrame())
    personnel_df = pd.DataFrame(people_df)

    def find_personnel_id(label):
        if personnel_df.empty:
            return None
        label = str(label).lower().replace("–", " ").replace("-", " ").strip()
        for _, r in personnel_df.iterrows():
            name_parts = " ".join(str(x).lower() for x in [r.get("Rank",""), r.get("FirstName",""), r.get("LastName","")] if str(x).strip())
            if name_parts and name_parts in label:
                return r.get("PersonnelID", None)
        return None

    inc_key = str(inc_num).strip()
    df = data.get("Incident_Personnel", pd.DataFrame())
    if df is None or df.empty:
        df = pd.DataFrame(columns=CHILD_TABLES["Incident_Personnel"])

    new_entries = []
    for n in picked_people:
        pid = find_personnel_id(n)
        new_entries.append({
            PRIMARY_KEY: inc_key,
            "Name": str(n).strip(),
            "Role": role_default,
            "Hours": hours_default,
            "RespondedIn": responded_in_default if responded_in_default else None,
            "Notes": None,
            "PersonnelID": pid
        })
    new_df = pd.concat([df, pd.DataFrame(new_entries)], ignore_index=True)
    data["Incident_Personnel"] = new_df
    return data
# --- END PATCH ---
