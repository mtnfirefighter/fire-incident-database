# patch_personnelid_v4_3_2.py
# Drop-in helper for fire-incident-database v4.3.2
# Purpose: when adding members to an incident, also store PersonnelID from the Personnel roster.

import pandas as pd

def build_rows_with_ids(picked_people, people_df, inc_key, role_default, hours_default, responded_in_default, PRIMARY_KEY):
    """Return list of row dicts for Incident_Personnel including PersonnelID.")
    """
    def _norm(s): 
        return " ".join(str(s or "").strip().lower().split())

    rows = []
    # Pre-build a quick search structure for roster lookups
    roster = people_df.copy() if people_df is not None else pd.DataFrame()
    if roster is None or roster.empty:
        roster = pd.DataFrame(columns=["PersonnelID","Name","FirstName","LastName","FullName","Rank"])

    # Make sure roster has Name column for display
    if "Name" not in roster.columns:
        fn = roster["FirstName"].astype(str) if "FirstName" in roster.columns else ""
        ln = roster["LastName"].astype(str) if "LastName" in roster.columns else ""
        roster["Name"] = (fn.str.strip() + " " + ln.str.strip()).str.strip()

    def _lookup_person_id_from_label(label):
        target = _norm(label)
        if roster.empty:
            return None, label
        for _, r in roster.iterrows():
            pid = r.get("PersonnelID", None)
            full = str(r.get("Name") or r.get("FullName") or "").strip()
            fn = str(r.get("FirstName") or "").strip()
            ln = str(r.get("LastName") or "").strip()
            rk = str(r.get("Rank") or "").strip()
            candidates = [
                full,
                f"{rk} {fn} {ln}".strip(),
                f"{fn} {ln}".strip(),
                f"{ln}, {fn}".strip(", "),
            ]
            for c in candidates:
                if _norm(c) == target and full:
                    return (None if pd.isna(pid) else str(pid)), full
        return None, label  # fallback

    for lbl in picked_people:
        pid, disp = _lookup_person_id_from_label(lbl)
        rows.append({
            PRIMARY_KEY: str(inc_key),
            "PersonnelID": pid,
            "Name": disp,
            "Role": role_default,
            "Hours": hours_default,
            "RespondedIn": (responded_in_default or None),
        })
    return rows