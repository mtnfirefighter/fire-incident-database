
# patch_personnelid_assign_v4_3_2.py
"""
Purpose
-------
Fill PersonnelID automatically when adding members to an incident.

How to use (two tiny edits in app.py)
-------------------------------------
1) Add this import near the top with your other imports:
       import patch_personnelid_assign_v4_3_2 as ppa

2) In the Write/Report tab where you handle the "Add Selected Members" button:
   Replace your manual append logic with:

       if st.button("Add Selected Members", key="btn_add_members"):
           new_rows = ppa.build_personnel_rows(
               incident_id=sel,                      # or your incident key variable
               picks=selected_members,               # your multiselect list of names
               default_role=default_role,            # your default role string
               default_hours=default_hours,          # float or str
               responded_in=responded_in,            # optional str (or None/"")
               notes=None,                           # optional notes field
               data=data                             # the in-memory data dict with "Personnel" & "Incident_Personnel"
           )
           data["Incident_Personnel"] = ppa.merge_into_incident_personnel(
               data.get("Incident_Personnel"), new_rows, primary_key=PRIMARY_KEY
           )
           st.success(f"Added {len(new_rows)} member(s) to incident {sel}.")

This patch ONLY affects how rows are added to Incident_Personnel. It does not
change your schema or other tabs. If something looks off, comment out the import
and the call above and your old behavior will return immediately.
"""
from __future__ import annotations
import pandas as pd
from typing import Iterable, List, Dict, Optional

def _norm(s: Optional[str]) -> str:
    if s is None:
        return ""
    return " ".join(str(s).strip().lower().split())

def _label_variants(row: pd.Series) -> List[str]:
    first = _norm(row.get("FirstName"))
    last  = _norm(row.get("LastName"))
    rank  = _norm(row.get("Rank"))
    # Common label variants used in UIs
    variants = [
        f"{first} {last}".strip(),
        f"{last}, {first}".strip(", "),
        f"{rank} {first} {last}".strip(),
        f"{first} {last} {rank}".strip(),
        f"{rank} {last}".strip(),
        f"{first}".strip(),
        f"{last}".strip(),
    ]
    # Deduplicate while preserving order
    seen, out = set(), []
    for v in variants:
        if v and v not in seen:
            out.append(v); seen.add(v)
    return out

def _build_roster_index(roster: pd.DataFrame) -> Dict[str, Dict[str, str]]:
    """
    Returns a lookup dict: normalized_label -> {"PersonnelID": "...", "Name": "First Last"}.
    Works with columns: PersonnelID, FirstName, LastName, Rank (free text allowed).
    """
    if roster is None or roster.empty:
        return {}
    r = roster.copy()

    # normalize column names if needed
    if "PersonnelID" not in r.columns:
        for alt in ("ID","MemberID"):
            if alt in r.columns:
                r = r.rename(columns={alt: "PersonnelID"})
                break
    if "FirstName" not in r.columns and "First" in r.columns:
        r = r.rename(columns={"First":"FirstName"})
    if "LastName" not in r.columns and "Last" in r.columns:
        r = r.rename(columns={"Last":"LastName"})

    idx: Dict[str, Dict[str, str]] = {}
    for _, row in r.iterrows():
        pid = row.get("PersonnelID")
        first = row.get("FirstName", "")
        last  = row.get("LastName", "")
        display_name = " ".join([str(first or "").strip(), str(last or "").strip()]).strip()
        for label in _label_variants(row):
            idx[_norm(label)] = {"PersonnelID": None if pd.isna(pid) else str(pid), "Name": display_name}
    return idx

def lookup_person_ids(data: dict, picks: Iterable[str]) -> List[Dict[str,str]]:
    """
    Map user-picked labels to {'PersonnelID','Name'} using the Personnel roster.
    """
    roster = data.get("Personnel", pd.DataFrame())
    idx = _build_roster_index(roster)
    out: List[Dict[str,str]] = []
    for p in (picks or []):
        l = _norm(p)
        rec = idx.get(l, {"PersonnelID": None, "Name": p})
        out.append(rec)
    return out

def build_personnel_rows(
    incident_id: str,
    picks: Iterable[str],
    default_role: str,
    default_hours,
    responded_in: Optional[str],
    notes: Optional[str],
    data: dict,
) -> pd.DataFrame:
    """
    Build a dataframe of rows for Incident_Personnel with PersonnelID filled.
    Columns: IncidentNumber, PersonnelID, Name (optional), Role, Hours, RespondedIn, Notes
    """
    matches = lookup_person_ids(data, picks)
    rows = []
    for m in matches:
        rows.append({
            "IncidentNumber": incident_id,
            "PersonnelID": m.get("PersonnelID"),
            "Name": m.get("Name"),
            "Role": default_role,
            "Hours": default_hours,
            "RespondedIn": responded_in if responded_in else None,
            "Notes": notes if notes else None,
        })
    return pd.DataFrame(rows)

def merge_into_incident_personnel(current: Optional[pd.DataFrame], new_rows: pd.DataFrame, primary_key: str) -> pd.DataFrame:
    """
    Append new rows and avoid exact duplicates by [IncidentNumber, PersonnelID, Role, Hours, RespondedIn].
    """
    if current is None or current.empty:
        base = pd.DataFrame(columns=["IncidentNumber","PersonnelID","Name","Role","Hours","RespondedIn","Notes"])
    else:
        base = current.copy()

    combined = pd.concat([base, new_rows], ignore_index=True)
    # drop exact duplicates on key fields (keep first)
    keep_cols = [c for c in ["IncidentNumber","PersonnelID","Name","Role","Hours","RespondedIn"] if c in combined.columns]
    if keep_cols:
        combined = combined.drop_duplicates(subset=keep_cols, keep="first").reset_index(drop=True)
    return combined
