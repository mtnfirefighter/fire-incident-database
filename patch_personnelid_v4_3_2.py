
# patch_personnelid_v4_3_2.py
# Minimal patch: provide a helper to render Personnel in Print with PersonnelID merged from roster.
# Does not modify your data; only affects the Print tab rendering.
import pandas as pd
import html as _html

def _esc(x):
    return _html.escape("" if x is None else str(x))

def _ensure_col(df, colname):
    if df is None:
        return pd.DataFrame({colname: []})
    if colname not in df.columns:
        df[colname] = pd.NA
    return df

def _normalize_roster(roster: pd.DataFrame) -> pd.DataFrame:
    if roster is None or roster.empty:
        return pd.DataFrame()
    r = roster.copy()
    if "PersonnelID" not in r.columns:
        for alt in ["ID", "MemberID"]:
            if alt in r.columns:
                r = r.rename(columns={alt: "PersonnelID"})
                break
    if "Name" not in r.columns:
        for alt in ["FullName", "MemberName"]:
            if alt in r.columns:
                r = r.rename(columns={alt: "Name"})
                break
    return r

def personnel_table_html(data: dict, base_df: pd.DataFrame, sel, PRIMARY_KEY: str, CHILD_TABLES: dict, ip_df: pd.DataFrame = None) -> str:
    """
    Returns an HTML table string for Personnel on Scene showing:
        PersonnelID, Name, Role, Hours, RespondedIn
    Falls back gracefully if columns are missing.
    """
    if not isinstance(data, dict) or "Incident_Personnel" not in data:
        # safest fallback: empty table
        return "<i>No personnel recorded</i>"

    if ip_df is None:
        ip_df = data.get("Incident_Personnel", pd.DataFrame())
    ip_df = _ensure_col(ip_df, "PersonnelID")
    if PRIMARY_KEY not in ip_df.columns:
        # Can't filter; show best effort
        subset = ip_df.copy()
    else:
        subset = ip_df[ip_df[PRIMARY_KEY].astype(str) == str(sel)].copy()

    roster = _normalize_roster(data.get("Personnel", pd.DataFrame()))
    if not subset.empty and not roster.empty and "Name" in subset.columns and "Name" in roster.columns and "PersonnelID" in roster.columns:
        merged = subset.merge(
            roster[["Name", "PersonnelID"]].drop_duplicates(),
            on="Name", how="left", suffixes=("", "_roster")
        )
        if "PersonnelID_roster" in merged.columns:
            merged["PersonnelID"] = merged["PersonnelID"].fillna(merged["PersonnelID_roster"])
            merged = merged.drop(columns=[c for c in ["PersonnelID_roster"] if c in merged.columns])
        df = merged
    else:
        df = subset

    show_cols = [c for c in ["PersonnelID","Name","Role","Hours","RespondedIn"] if c in df.columns]
    if not show_cols:
        show_cols = list(df.columns)

    try:
        return df[show_cols].to_html(index=False)
    except Exception:
        return df.to_html(index=False)
